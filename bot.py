"""bot.py

Telegram disc golf bot — utilities and Telegram bot handlers for
managing tournament folders and interacting with users via a
Telegram bot.

Author: Baidiuk Anton DA-22

This module is prepared for automatic documentation generation with
Sphinx/autodoc. Long-running or environment-specific operations are
guarded so importing the module during documentation build does not
execute those operations.

Public objects include ``initialize_log_file``, ``log_usage``,
``build_image_path`` and the Telegram handler ``start``.
"""

# Google Colab utilities might not be available in all environments
try:
    from google.colab import files
except Exception:
    files = None

import os
# Імпорт модулів для роботи з Telegram API та асинхронними задачами
from telegram import Update, ReplyKeyboardMarkup, InlineKeyboardMarkup, InlineKeyboardButton
from telegram.ext import ApplicationBuilder, CommandHandler, MessageHandler, filters, ContextTypes, ConversationHandler, CallbackQueryHandler, ContextTypes

# Додаткові бібліотеки
try:
    import nest_asyncio
except Exception:
    nest_asyncio = None
import pandas as pd
import warnings
import numpy as np
import matplotlib.pyplot as plt
import matplotlib.cm as cm
try:
    from google.colab import drive
except Exception:
    drive = None
from IPython.display import display
import asyncio
TOKEN = "7590977209:AAHbx8cPnmH6TRKWXjKAOXhb8TX-jqFHMDY"
main_folder = "tournaments"

# Guard side-effects so Sphinx (or other importers) can import the
# module without executing long-running or environment-specific code.
if __name__ == "__main__":
    # Mount Google Drive (only when running as a script)
    try:
        drive.mount('/content/drive')
    except Exception:
        pass

    # Apply nest_asyncio (only when running as a script)
    try:
        nest_asyncio.apply()
    except Exception:
        pass

    # Try to load external PDGA data if available
    try:
        pdga_df = pd.read_excel('/content/drive/MyDrive/dg/PDGA.xlsx')
    except Exception:
        pdga_df = None



import os
import shutil

# Головна папка для збереження турнірів
main_folder = "tournaments"

# The code that creates or removes folders is guarded so importing the
# module for documentation does not cause filesystem changes. Run the
# block below only when executing the script directly.
if __name__ == "__main__":
    # Якщо головна папка існує — видаляємо її повністю
    if os.path.exists(main_folder):
        shutil.rmtree(main_folder)

    # Визначення структури турнірів
    structure = {
        # "2024": ["Lutsk Open"],  # приклад для попереднього року
        "2025": [
            "Kyiv Open",
            "Lutsk Open",
            "Kremenchuk Open",
            "Frozen Chainz",
            "The First Spring Championship",
            "Ukrainian Cup",
            "Lubart Cup",
            "Kremenchuk Autumn Cup"
        ]
    }

    # Створення папок відповідно до структури
    for year, tournaments in structure.items():
        year_path = os.path.join(main_folder, year)
        os.makedirs(year_path, exist_ok=True)
        for tournament in tournaments:
            tournament_path = os.path.join(year_path, tournament)
            os.makedirs(tournament_path, exist_ok=True)

# Manual creation of divisions and player folders for local runs only.
# This block is executed only when running the script directly and is
# skipped during import (e.g., when Sphinx autodoc imports the module).
if __name__ == "__main__":
    # Ручно задані дивізіони
    MPO = os.path.join("tournaments", "2024", "Lutsk Open", "MPO")
    MA4 = os.path.join("tournaments", "2024", "Lutsk Open", "MA4")
    FA1 = os.path.join("tournaments", "2024", "Lutsk Open", "FA1")
    MA40 = os.path.join("tournaments", "2024", "Lutsk Open", "MA40")

    # Створення папок (локально)
    os.makedirs(MPO, exist_ok=True)
    os.makedirs(MA4, exist_ok=True)
    os.makedirs(FA1, exist_ok=True)
    os.makedirs(MA40, exist_ok=True)

    base_path1 = os.path.join("tournaments", "2024", "Lutsk Open", "MPO")
    player_names = [
        "Bohdan Aleksieiev",
        "Oleksandr Ratushniak",
        "Oleksandr Romanenko",
        "Mykola Korsai",
        "Ruvym Muts",
        "Oleh Mykhalevych",
        "Artur Zhuravliov",
        "Dima Boiko"
    ]

    for name in player_names:
        os.makedirs(os.path.join(base_path1, name), exist_ok=True)

    base_path2 = os.path.join("tournaments", "2024", "Lutsk Open", "MA40")
    player_names = [
        "Serhii Turchyk",
        "Charles Czepyha",
        "Andrii Legin",
        "Ruslan Muts"
    ]

    for name in player_names:
        os.makedirs(os.path.join(base_path2, name), exist_ok=True)

    base_path3 = os.path.join("tournaments", "2024", "Lutsk Open", "FA1")
    player_names = [
        "Ira Sheremeta",
        "Inna Pashchenko",
        "Anna Romanenko",
        "Viktoria Muts"
    ]

    for name in player_names:
        os.makedirs(os.path.join(base_path3, name), exist_ok=True)

    base_path4 = os.path.join("tournaments", "2024", "Lutsk Open", "MA4")
    player_names = [
        "Oleksii Pashchenko",
        "Yurii Pliekhotkin",
        "Illia Kravchenko",
        "Vadym Novikov",
        "Andrii Hnidets",
        "Anton Baidiuk"
    ]

    for name in player_names:
        os.makedirs(os.path.join(base_path4, name), exist_ok=True)

"""
Модуль для ведення журналу використання Telegram‑бота.

Функціонал:
- створює файл журналу, якщо він відсутній;
- записує інформацію про кожну дію користувача (chat_id, ім'я, час, дія);
- зберігає дані у форматі Excel.
"""

import os
import pandas as pd
from datetime import datetime

LOG_PATH = "/content/drive/MyDrive/dg/usage_log.xlsx"

def initialize_log_file():
    """
    Ініціалізує файл журналу використання.

    Якщо файл журналу не існує, створює новий Excel‑файл
    з колонками: chat_id, user_name, timestamp, action.
    """
    if not os.path.exists(LOG_PATH):
        df = pd.DataFrame(columns=["chat_id", "user_name", "timestamp", "action"])
        df.to_excel(LOG_PATH, index=False)

def log_usage(chat_id: int, user_name: str, action: str):
    """
    Записує новий запис у журнал використання.

    :param chat_id: унікальний ідентифікатор чату користувача
    :type chat_id: int
    :param user_name: ім'я користувача
    :type user_name: str
    :param action: виконана дія (наприклад, команда бота)
    :type action: str
    :return: None
    :rtype: None
    
    Алгоритм:

    - Перевіряє наявність файлу журналу, створює його при потребі.
    - Читає існуючі дані з Excel.
    - Додає новий рядок із даними користувача та дією.
    - Зберігає оновлений журнал у Excel.
    """
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    if not os.path.exists(LOG_PATH):
        initialize_log_file()
    df = pd.read_excel(LOG_PATH)
    new_row = {
        "chat_id": chat_id,
        "user_name": user_name,
        "timestamp": timestamp,
        "action": action
    }
    df = pd.concat([df, pd.DataFrame([new_row])], ignore_index=True)
    df.to_excel(LOG_PATH, index=False)
def build_image_path(base_dir: str, year: str, tournament: str, division: str, suffix: str, player: str = None) -> str:
    """
    Формує шлях до зображення для турніру.

    :param base_dir: базова директорія (наприклад, "/base")
    :type base_dir: str
    :param year: рік турніру (наприклад, "2025")
    :type year: str
    :param tournament: назва турніру (наприклад, "Ukrainian Cup")
    :type tournament: str
    :param division: дивізіон (наприклад, "MPO")
    :type division: str
    :param suffix: суфікс файлу (наприклад, "best", "ideal")
    :type suffix: str
    :param player: ім’я гравця (опціонально, якщо потрібно створити шлях всередині папки гравця)
    :type player: str, optional
    :return: сформований шлях до файлу .png
    :rtype: str

    Приклади:
    >>> build_image_path("/base", "2023", "OpenCup", "MA1", "best")
    '/base/2023/OpenCup/MA1/best.png'

    >>> build_image_path("/base", "2023", "OpenCup", "MA1", "ideal", "JohnDoe")
    '/base/2023/OpenCup/MA1/JohnDoe/ideal.png'
    """
    if player:
        return os.path.join(base_dir, year, tournament, division, player, f"{suffix}.png")
    return os.path.join(base_dir, year, tournament, division, f"{suffix}.png")

# Модуль із функцією старту Telegram‑бота.
#
# Функціонал:
# - очищає історію повідомлень користувача;
# - видаляє попередні системні повідомлення та повідомлення виклику;
# - створює нове системне повідомлення з кнопкою /start;
# - формує головне меню з опціями: Турнір, Гравець, Інструкція;
# - логує використання команди /start у журналі.
BASE_DIR = main_folder
# Do not read external files at import time; load PDGA data when running as script
pdga_df = None

from telegram import Update, ReplyKeyboardMarkup, InlineKeyboardMarkup, InlineKeyboardButton
from telegram.ext import ApplicationBuilder, CommandHandler, MessageHandler, filters, ContextTypes, ConversationHandler, CallbackQueryHandler, ContextTypes
async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """
    Обробник команди /start для Telegram‑бота.

    :param update: об’єкт Update, що містить інформацію про повідомлення та користувача
    :type update: telegram.Update
    :param context: контекст виконання, що містить дані користувача та методи для взаємодії з ботом
    :type context: telegram.ext.ContextTypes.DEFAULT_TYPE
    :return: None
    :rtype: None

    Алгоритм:

    1. Логує використання команди /start.
    2. Очищає історію повідомлень користувача.
    3. Видаляє повідомлення виклику та попереднє системне повідомлення.
    4. Створює нове системне повідомлення з кнопкою /start.
    5. Формує головне меню з кнопками:

       - 📊 Турнір
       - 👤 Гравець
       - ℹ️ Інструкція

    6. Зберігає ідентифікатори створених повідомлень у контексті користувача.
    """
    chat_id = update.effective_chat.id
    user_name = update.effective_user.full_name

    log_usage(chat_id, user_name, "Команда /start")
    # 🧹 Очистка історії
    for msg_id in context.user_data.get("history", []):
        try:
            await context.bot.delete_message(chat_id=chat_id, message_id=msg_id)
        except:
            pass
    context.user_data["history"] = []

    # 🗑️ Видалення повідомлення виклику
    if update.message:
        try:
            await context.bot.delete_message(chat_id=chat_id, message_id=update.message.message_id)
        except:
            pass

    # 🧽 Видалити попереднє системне повідомлення, якщо є
    reply_msg_id = context.user_data.get("reply_keyboard_msg_id")
    if reply_msg_id:
        try:
            await context.bot.delete_message(chat_id=chat_id, message_id=reply_msg_id)
        except:
            pass

    # 🆕 Створити нове системне повідомлення
    msg_reply = await context.bot.send_message(
        chat_id=chat_id,
        text="⌨️ Якщо щось піде не так — натисни /start",
        reply_markup=ReplyKeyboardMarkup([["/start"]], resize_keyboard=True)
    )
    context.user_data["reply_keyboard_msg_id"] = msg_reply.message_id

    # 📋 Головне меню
    keyboard = [
        [InlineKeyboardButton("📊 Турнір", callback_data="menu:tournament")],
        [InlineKeyboardButton("👤 Гравець", callback_data="menu:player")],
        [InlineKeyboardButton("ℹ️ Інструкція", callback_data="menu:instructions")]
    ]
    reply_markup = InlineKeyboardMarkup(keyboard)
    msg_menu = await context.bot.send_message(
        chat_id=chat_id,
        text="Обери опцію:",
        reply_markup=reply_markup
    )
    context.user_data["history"].append(msg_menu.message_id)

async def send_round_table(query, context, round_type, caption_text):
    """
    Відправляє таблицю результатів раунду у Telegram‑чат.

    :param query: об’єкт запиту, що містить повідомлення та дані виклику
    :type query: telegram.CallbackQuery
    :param context: контекст виконання, що містить дані користувача та методи для взаємодії з ботом
    :type context: telegram.ext.ContextTypes.DEFAULT_TYPE
    :param round_type: тип раунду (наприклад, "ideal_round", "horror_round", "best_round", "worst_round")
    :type round_type: str
    :param caption_text: текст підпису до зображення, може містити форматовані змінні {tournament}, {year}, {division}
    :type caption_text: str
    :return: None
    :rtype: None

     Алгоритм:

     - Очищає історію повідомлень користувача та видаляє повідомлення виклику.
     - Отримує параметри турніру (рік, назву, дивізіон, гравця).
     - Перевіряє, чи задано турнір та рік; якщо ні — повідомляє про помилку.
     - Визначає суфікс файлу для зображення залежно від типу раунду.
     - Формує шлях до файлу зображення:

         - якщо задано гравця → шлях включає його ім’я;
         - якщо гравець не заданий → шлях формується лише до дивізіону.

     - Відправляє фото з підписом у чат.
     - Формує меню з опціями:

         - 🥏 Ідеальний раунд
         - 💩 Жахливий раунд
         - 🌟 Найкращий раунд
         - 🕳️ Найгірший раунд
         - 📈 Успішність на кошиках (якщо вибрано гравця)
         - 🔙 Назад до гравців (дивізіон)
         - 🏠 Головне меню

     - Якщо файл не знайдено — повідомляє про помилку.
    """
    chat_id = query.message.chat_id

    # 🧹 Очистити історію
    for msg_id in context.user_data.get("history", []):
        try:
            await context.bot.delete_message(chat_id=chat_id, message_id=msg_id)
        except:
            pass
    context.user_data["history"] = []

    # 🗑️ Видалити повідомлення виклику
    try:
        await context.bot.delete_message(chat_id=chat_id, message_id=query.message.message_id)
    except:
        pass

    # 📥 Параметри
    year = context.user_data.get("year", "")
    tournament = context.user_data.get("tournament", "")
    division = context.user_data.get("division", "ALL")
    player = context.user_data.get("player", "")  # ← буде порожнім, якщо запит йде з дивізіону

    if not tournament or not year:
        msg = await context.bot.send_message(chat_id=chat_id, text="⚠️ Не вказано турнір або рік.")
        context.user_data["history"].append(msg.message_id)
        return

    suffix_map = {
        "ideal_round": "ideal",
        "horror_round": "horror",
        "best_round": "best",
        "worst_round": "worst"
    }
    suffix = suffix_map.get(round_type, round_type)

    # 🔍 Формуємо шлях
    if player:
        file_path = f"/content/drive/MyDrive/dg/{year}/{tournament}/{division}/{player}/{suffix}.png"
    else:
        file_path = f"/content/drive/MyDrive/dg/{year}/{tournament}/{division}/{suffix}.png"

    try:
        with open(file_path, "rb") as img:
            msg = await context.bot.send_photo(
                chat_id=chat_id,
                photo=img,
                caption=caption_text.format(tournament=tournament, year=year, division=division),
                parse_mode="Markdown"
            )
            context.user_data["history"].append(msg.message_id)

        # 🔁 Меню після фото
        menu_text = context.user_data.get("previous_menu_text", "⬇️ Обери наступну опцію:")
        keyboard = [
            [InlineKeyboardButton("🥏 Ідеальний раунд", callback_data="ideal_round")],
            [InlineKeyboardButton("💩 Жахливий раунд", callback_data="horror_round")],
            [InlineKeyboardButton("🌟 Найкращий раунд", callback_data="best_round")],
            [InlineKeyboardButton("🕳️ Найгірший раунд", callback_data="worst_round")]

        ]
        if player is not None and division != "ALL":
            keyboard.append([InlineKeyboardButton("📈 Успішність на кошиках", callback_data="player:percent")])


        # 🔙 Назад — гравці або дивізіон
        if player:
            keyboard.append([InlineKeyboardButton(f"🔙 Назад до гравців ({division})", callback_data=f"division:{division}")])
        else:
            keyboard.append([InlineKeyboardButton(f"🔙 Назад до гравців ({division})", callback_data=f"division:{division}")])

        keyboard.append([InlineKeyboardButton("🏠 Головне меню", callback_data="back:main")])

        msg = await context.bot.send_message(
            chat_id=chat_id,
            text=menu_text,
            reply_markup=InlineKeyboardMarkup(keyboard)
        )
        context.user_data["history"].append(msg.message_id)

    except FileNotFoundError:
        msg = await context.bot.send_message(
            chat_id=chat_id,
            text=f"❌ Файл `{file_path}` не знайдено.",
            parse_mode="Markdown"
        )
        context.user_data["history"].append(msg.message_id)


async def send_image_with_caption(query, context, file_path, caption, followup_keyboard=None):
    """
    Надсилає зображення з підписом у Telegram‑чат та опціонально додає меню.

    :param query: об’єкт запиту, що містить повідомлення та дані виклику
    :type query: telegram.CallbackQuery
    :param context: контекст виконання, що містить дані користувача та методи для взаємодії з ботом
    :type context: telegram.ext.ContextTypes.DEFAULT_TYPE
    :param file_path: шлях до файлу зображення
    :type file_path: str
    :param caption: текст підпису до зображення
    :type caption: str
    :param followup_keyboard: клавіатура для наступного меню (опціонально)
    :type followup_keyboard: list[list[telegram.InlineKeyboardButton]], optional
    :return: None
    :rtype: None
     Алгоритм:

     1. Очищає історію повідомлень користувача.
     2. Видаляє повідомлення кнопки, що викликало функцію.
     3. Відправляє зображення з підписом:

         - якщо файл існує → надсилає фото;
         - якщо файл не знайдено → повідомляє про помилку.

     4. Якщо передано клавіатуру → надсилає меню з опціями.
     5. Зберігає ідентифікатори створених повідомлень у контексті користувача.
    """
    chat_id = query.message.chat_id

    # 🧹 Очистка історії повідомлень
    for msg_id in context.user_data.get("history", []):
        try:
            await context.bot.delete_message(chat_id=chat_id, message_id=msg_id)
        except:
            pass
    context.user_data["history"] = []

    # 🗑️ Видалити повідомлення кнопки, якщо існує
    try:
        await context.bot.delete_message(chat_id=chat_id, message_id=query.message.message_id)
    except:
        pass

    # 📸 Надсилання зображення
    try:
        with open(file_path, "rb") as img:
            photo_msg = await context.bot.send_photo(
                chat_id=chat_id,
                photo=img,
                caption=caption
            )
            context.user_data["history"].append(photo_msg.message_id)
    except FileNotFoundError:
        error_msg = await context.bot.send_message(
            chat_id=chat_id,
            text=f"❌ Файл `{file_path}` не знайдено.",
            parse_mode="Markdown"
        )
        context.user_data["history"].append(error_msg.message_id)
        return

    # ⏭️ Надсилання меню (якщо передано)
    if followup_keyboard:
        menu_msg = await context.bot.send_message(
            chat_id=chat_id,
            text="⬇️ Обери наступну опцію:",
            reply_markup=InlineKeyboardMarkup(followup_keyboard)
        )
        context.user_data["history"].append(menu_msg.message_id)


async def show_basket_selection(query, context):
    """
    Відображає меню вибору кошика у Telegram‑чаті.

    :param query: об’єкт запиту, що містить повідомлення та дані виклику
    :type query: telegram.CallbackQuery
    :param context: контекст виконання, що містить дані користувача та методи для взаємодії з ботом
    :type context: telegram.ext.ContextTypes.DEFAULT_TYPE
    :return: None
    :rtype: None
     Алгоритм:

     1. Формує кнопки з номерами кошиків від 1 до 18.
     2. Розбиває кнопки на рядки по 6 елементів для зручності перегляду.
     3. Додає нижній рядок із кнопками:

         - 🔙 Назад (повернення до меню кошиків)
         - 🏠 Головне меню (повернення до головного меню).

     4. Оновлює повідомлення у чаті, відображаючи меню вибору кошика.
    """
    basket_buttons = [
        InlineKeyboardButton(str(i), callback_data=f"basket:{i}")
        for i in range(1, 19)
    ]
    keyboard = [basket_buttons[i:i+6] for i in range(0, len(basket_buttons), 6)]
    keyboard.append([
        InlineKeyboardButton("🔙 Назад", callback_data="menu:baskets"),
        InlineKeyboardButton("🏠 Головне меню", callback_data="back:main")
    ])
    await query.edit_message_text(
        "🔍 Обери номер кошика для перегляду:",
        reply_markup=InlineKeyboardMarkup(keyboard)
    )

# Обробка кнопок
async def button_handler(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """
    Центральний обробник callback‑кнопок Telegram‑бота.

    :param update: об’єкт Update, що містить callback_query та дані повідомлення
    :type update: telegram.Update
    :param context: контекст виконання, що містить user_data та методи для взаємодії з ботом
    :type context: telegram.ext.ContextTypes.DEFAULT_TYPE
    :return: None
    :rtype: None

     Основні сценарії:

     - ``menu:instructions``: показує інструкцію користування ботом.
     - ``menu:tournament``: відкриває список років турнірів.
     - ``year:YYYY``: показує турніри за вибраний рік.
     - ``tournament:NAME``: відкриває дивізіони турніру.
     - ``division:DIV``: показує список гравців у дивізіоні.
     - ``player:NAME``: відкриває меню статистики для конкретного гравця.
     - ``division:ALL``: показує загальну статистику турніру.
     - ``send_table``: надсилає турнірну таблицю.
     - ``menu:baskets``: відкриває меню аналізу кошиків.
     - ``baskets:top``, ``baskets:hardest``, ``baskets:overview``, ``basket:N``: надсилає відповідні графіки по кошиках.
     - ``menu:stability``: надсилає графік стабільності.
     - ``ideal_round``, ``horror_round``, ``best_round``, ``worst_round``: надсилає таблиці результатів раундів.
     - ``player:percent``: показує успішність гравця на кошиках.
     - ``restart:start``: перезапускає логіку ``/start``.
     - ``back:main``: очищає історію та повертає головне меню.

     Логіка:

     1. Логує натискання кнопки.
     2. В залежності від ``callback_data`` виконує потрібний сценарій.
     3. Використовує допоміжні функції:

         - ``send_image_with_caption``
         - ``send_round_table``
         - ``show_basket_selection``
         - ``start``

     4. Оновлює ``context.user_data["history"]`` для контролю повідомлень.
    """
    query = update.callback_query
    await query.answer()
    data = query.data
    chat_id = query.message.chat_id
    user_name = query.from_user.full_name
    log_usage(chat_id, user_name, f"Натиснута кнопка: {data}")

    app.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND, handle_text))

    if data == "menu:instructions":
        instruction_text = (
            "📌 *Інструкція користування ботом*\n\n"
            "Поточні гілки бота:\n"
            "`Турнір → Рік → Турнір → Дивізіон → Гравець → Статистика`\n"
            "`Гравець → Турніри гравця → Персональна статистика`\n\n"
            "Пояснення статистики:\n"
            "🥏 *Ідеальний раунд*: таблиця із найкращих результатів на кожному кошику\n"
            "💩 *Жахливий раунд*: таблиця із найгірших результатів на кожному кошику\n"
            "🌟 *Найкращий раунд*: таблиця реальних ігрових раундів із найкращим результатом\n"
            "🕳️ *Найгірший раунд*: таблиця реальних ігрових раундів із найгіршим результатом\n"
            "🧱 *Стабільність*: показує чи стабільно гравець відігравав на кожному кошику\n"
            "   іншими словами, наскільки випадкові результати гравця\n"
            "📈 *Успішність на кошиках*: показує відсоткову різницю між `par` та середнім результатом\n"
            "🧺 *По кошиках*:\n"
            "🏅 *Кращі гравці на кошиках*: топ середніх результатів гравців\n"
            "🔥 *Складність кошиків*: кошики посортовані по середньому результату\n"
            "📊 *Загальний огляд кошиків*: середні значення та результати\n\n"
            "Надсилайте баги, помилки та інші неточності для покращення! \n"
            "💡 Усі пропозиції та ідеї буде враховано \n"
            " Зворотній зв'язок — @antonbaidiuk"
        )

        keyboard = [
            [InlineKeyboardButton("🔙 Назад до меню", callback_data="restart:start")]
        ]
        reply_markup = InlineKeyboardMarkup(keyboard)

        edited = await query.edit_message_text(
            text=instruction_text,
            parse_mode="Markdown",
            reply_markup=reply_markup
        )
        context.user_data.setdefault("history", []).append(edited.message_id)

    elif data == "baskets:top":
         await show_basket_selection(query, context)


    elif data == "restart:start":
        await start(update, context)

    elif data == "player:percent":
        year = context.user_data.get("year")
        tournament = context.user_data.get("tournament")
        division = context.user_data.get("division")
        player = context.user_data.get("player")
        file_path = f"/content/drive/MyDrive/dg/{year}/{tournament}/{division}/{player}/percent.png"
        caption = f"📈 Успішність {player} на кожному кошику — {tournament} {year}, дивізіон: {division}"

        await send_image_with_caption(query, context, file_path, caption)

        # 🔁 Повторити кнопки як у гілці player:...
        keyboard = [
            [InlineKeyboardButton("🥏 Ідеальний раунд", callback_data="ideal_round")],
            [InlineKeyboardButton("💩 Жахливий раунд", callback_data="horror_round")],
            [InlineKeyboardButton("🌟 Найкращий раунд", callback_data="best_round")],
            [InlineKeyboardButton("🕳️ Найгірший раунд", callback_data="worst_round")],
            [InlineKeyboardButton("📈 Успішність на кошиках", callback_data="player:percent")],
            [InlineKeyboardButton(f"🔙 Назад до гравців ({division})", callback_data=f"division:{division}")],
            [InlineKeyboardButton("🏠 Головне меню", callback_data="back:main")]
        ]


        msg = await context.bot.send_message(
            chat_id=query.message.chat_id,
            text=f"⬇️ Обери наступну опцію для {player} :",
            reply_markup=InlineKeyboardMarkup(keyboard)
        )
        context.user_data["history"].append(msg.message_id)


    elif data == "menu:tournament":
        years = [d for d in os.listdir(BASE_DIR) if os.path.isdir(os.path.join(BASE_DIR, d))]
        keyboard = [[InlineKeyboardButton(year, callback_data=f"year:{year}")] for year in sorted(years)]
        keyboard.append([InlineKeyboardButton("🏠 Головне меню", callback_data="back:main")])
        await query.edit_message_text("📆 Обери рік турніру:", reply_markup=InlineKeyboardMarkup(keyboard))

    elif data.startswith("year:") or data.startswith("back:year:"):
        if data.startswith("year:"):
          selected_year = data.split(":")[1]
        elif data.startswith("back:year:"):
          selected_year = data.split(":")[2]
        context.user_data["year"] = selected_year
        path = os.path.join(BASE_DIR, selected_year)

        if not os.path.exists(path):
          await query.edit_message_text(
              text=f"❌ Директорія для року `{selected_year}` не знайдена.",
              parse_mode="Markdown"
          )
          return

        tournaments = [d for d in os.listdir(path) if os.path.isdir(os.path.join(path, d))]

        keyboard = [[InlineKeyboardButton(t, callback_data=f"tournament:{t}")] for t in sorted(tournaments)]
        keyboard.append([InlineKeyboardButton("🔙 Назад до років", callback_data="menu:tournament")])
        keyboard.append([InlineKeyboardButton("🏠 Головне меню", callback_data="back:main")])

        await query.edit_message_text(
            f"📋 Турніри за {selected_year}:",
            reply_markup=InlineKeyboardMarkup(keyboard)
        )


    elif data.startswith("player:"):
        player = data.split(":")[1]
        context.user_data["player"] = player
        year = context.user_data["year"]
        tournament = context.user_data["tournament"]
        division = context.user_data["division"]

        keyboard = [
            [InlineKeyboardButton("🥏 Ідеальний раунд", callback_data="ideal_round")],
            [InlineKeyboardButton("💩 Жахливий раунд", callback_data="horror_round")],
            [InlineKeyboardButton("🌟 Найкращий раунд", callback_data="best_round")],
            [InlineKeyboardButton("🕳️ Найгірший раунд", callback_data="worst_round")],
        ]
        if player is not None and division != "ALL":
            keyboard.append([InlineKeyboardButton("📈 Успішність на кошиках", callback_data="player:percent")])


        keyboard.extend([
            [InlineKeyboardButton(f"🔙 Назад до гравців ({division})", callback_data=f"division:{division}")],
            [InlineKeyboardButton("🏠 Головне меню", callback_data="back:main")]
        ])


        await query.edit_message_text(
            f"✅ Обрано: {player} ({division}, {tournament} {year})\n\nОберіть опцію:",
            reply_markup=InlineKeyboardMarkup(keyboard)
        )


    elif data.startswith("tournament:") or data.startswith("back:tournament:"):
      parts = data.split(":")
      tournament = parts[1] if data.startswith("tournament:") else parts[2]

      # Якщо рік ще не заданий — зупиняємось
      year = context.user_data.get("year")
      if not year:
          await query.edit_message_text("⚠️ Переапустіть бота, будь ласка /start")
          return

      context.user_data["tournament"] = tournament
      path = os.path.join(BASE_DIR, year, tournament)
      divisions = [d for d in os.listdir(path) if os.path.isdir(os.path.join(path, d))]

      keyboard = [[InlineKeyboardButton("📦 Загальна статистика", callback_data="division:ALL")]]
      keyboard += [[InlineKeyboardButton(d, callback_data=f"division:{d}")] for d in sorted(divisions)]

      match_row = pdga_df[
          (pdga_df["event_year"] == int(year)) &
          (pdga_df["event_name"] == tournament)
      ]
      if not match_row.empty:
          pdga_url = match_row.iloc[0]["event_link"]
          keyboard.append([InlineKeyboardButton("🌐 Результат на PDGA", url=pdga_url)])

      keyboard.append([InlineKeyboardButton("🔙 Назад до турнірів", callback_data=f"back:year:{year}")])
      keyboard.append([InlineKeyboardButton("🏠 Головне меню", callback_data="back:main")])

      await query.edit_message_text(
          f"📁 Дивізіони в {tournament}:",
          reply_markup=InlineKeyboardMarkup(keyboard)
      )

    elif data == "division:ALL":
        context.user_data["player"] = ""
        context.user_data["division"] = "ALL"
        year = context.user_data["year"]
        tournament = context.user_data["tournament"]

        keyboard = [
            [InlineKeyboardButton("📋 Турнірна таблиця", callback_data="send_table")],
            [InlineKeyboardButton("🧺 По кошиках", callback_data="menu:baskets")],
            [InlineKeyboardButton("🧱 Стабільність", callback_data="menu:stability")],
            [InlineKeyboardButton("🥏 Ідеальний раунд", callback_data="ideal_round")],
            [InlineKeyboardButton("💩 Жахливий раунд", callback_data="horror_round")],
            [InlineKeyboardButton("🌟 Найкращий раунд", callback_data="best_round")],
            [InlineKeyboardButton("🕳️ Найгірший раунд", callback_data="worst_round")],
            [InlineKeyboardButton(f"🔙 Назад до дивізіонів {tournament} {year}", callback_data=f"back:tournament:{tournament}")],
            [InlineKeyboardButton("🏠 Main Menu", callback_data="back:main")]
        ]

        await query.edit_message_text(
            f"📦 Загальна статистика — турнір {tournament} {year}\n\nОбери опцію:",
            reply_markup=InlineKeyboardMarkup(keyboard)
        )
    elif data == "send_table":
        chat_id = query.message.chat_id
        year = context.user_data.get("year")
        tournament = context.user_data.get("tournament")
        division = context.user_data.get("division", "ALL")
        file_path = f"/content/drive/MyDrive/dg/{year}/{tournament}/{division}/table.png"

        # 🧹 Очистка
        for msg_id in context.user_data.get("history", []):
            try:
                await context.bot.delete_message(chat_id=chat_id, message_id=msg_id)
            except:
                pass
        context.user_data["history"] = []

        try:
            with open(file_path, "rb") as img:
                msg = await context.bot.send_photo(
                    chat_id=chat_id,
                    photo=img,
                    caption=f"📋 Турнірна таблиця — {tournament} {year}, дивізіон: {division}"
                )
                context.user_data["history"].append(msg.message_id)

            # 🔁 Повторне меню
            keyboard = [
                [InlineKeyboardButton("🧺 По кошиках", callback_data="menu:baskets")],
                [InlineKeyboardButton("🧱 Стабільність", callback_data="menu:stability")],
                [InlineKeyboardButton("📋 Турнірна таблиця", callback_data="send_table")],
                [InlineKeyboardButton("🥏 Ідеальний раунд", callback_data="ideal_round")],
                [InlineKeyboardButton("💩 Жахливий раунд", callback_data="horror_round")],
                [InlineKeyboardButton("🌟 Найкращий раунд", callback_data="best_round")],
                [InlineKeyboardButton("🕳️ Найгірший раунд", callback_data="worst_round")],
                [InlineKeyboardButton(f"🔙 Назад до дивізіонів {tournament} {year}", callback_data=f"back:tournament:{tournament}")],
                [InlineKeyboardButton("🏠 Main Menu", callback_data="back:main")]
            ]
            msg_menu = await context.bot.send_message(
                chat_id=chat_id,
                text="⬇️ Обери наступну опцію:",
                reply_markup=InlineKeyboardMarkup(keyboard)
            )
            context.user_data["history"].append(msg_menu.message_id)

        except FileNotFoundError:
            msg = await context.bot.send_message(
                chat_id=chat_id,
                text=f"❌ Файл `{file_path}` не знайдено.",
                parse_mode="Markdown"
            )
            context.user_data["history"].append(msg.message_id)

    elif data == "menu:baskets":
        chat_id = query.message.chat_id
        division = context.user_data.get("division", "ALL")
        tournament = context.user_data.get("tournament")
        year = context.user_data.get("year")

        # 🧹 Очистка історії
        for msg_id in context.user_data.get("history", []):
            try:
                await context.bot.delete_message(chat_id=chat_id, message_id=msg_id)
            except:
                pass
        context.user_data["history"] = []

        # 🧵 Кнопки для перегляду кошиків
        buttons = [
            [InlineKeyboardButton("🏅 Кращі гравці на кошиках", callback_data="baskets:top")],
            [InlineKeyboardButton("🔥 Складність кошиків", callback_data="baskets:hardest")],
            [InlineKeyboardButton("📊 Загальний огляд кошиків", callback_data="baskets:overview")]
        ]

        if division == "ALL":
            #buttons.append([InlineKeyboardButton("⚖️ Порівняння кошиків по дивізіонах", callback_data="baskets:comparison")])
            back_text = f"🔙 Назад до загальної статистики {tournament} {year}"
            back_callback = "menu:division_stat"
        else:
            back_text = f"🔙 Назад до статистики {division} {tournament} {year}"
            back_callback = "menu:division_stat"

        buttons.append([InlineKeyboardButton(back_text, callback_data=back_callback)])
        buttons.append([InlineKeyboardButton("🏠 Головне меню", callback_data="back:main")])

        msg = await context.bot.send_message(
            chat_id=chat_id,
            text=f"🧺 Аналіз кошиків — {tournament} {year}, дивізіон: {division}\n\nОберіть опцію:",
            reply_markup=InlineKeyboardMarkup(buttons)
        )
        context.user_data["history"].append(msg.message_id)
        try:
            await context.bot.delete_message(chat_id=chat_id, message_id=query.message.message_id)
        except:
            pass


    elif data == "baskets:top":
        # 🔄 Викликає поточну логіку з кнопками 1-18
        await query.answer()
        await button_handler(update, context)  # можеш винести логіку в окрему функцію при потребі

    elif data == "baskets:hardest":
        year = context.user_data.get("year")
        tournament = context.user_data.get("tournament")
        division = context.user_data.get("division")
        file_path = f"/content/drive/MyDrive/dg/{year}/{tournament}/{division}/holes.png"
        caption = f"🔥 Складність кошиків для дивізіону {division} — турнір {tournament} {year}"

        keyboard = [
            [InlineKeyboardButton("🏅 Кращі гравці на кошиках", callback_data="baskets:top")],
            [InlineKeyboardButton("🔥 Складність кошиків", callback_data="baskets:hardest")],
            [InlineKeyboardButton("📊 Загальний огляд кошиків", callback_data="baskets:overview")],
            [InlineKeyboardButton(f"🔙 Назад до статистики {division}", callback_data="menu:division_stat")],
            [InlineKeyboardButton("🏠 Головне меню", callback_data="back:main")]
        ]

        await send_image_with_caption(query, context, file_path, caption, followup_keyboard=keyboard)

    elif data == "baskets:overview":
        year = context.user_data.get("year")
        tournament = context.user_data.get("tournament")
        division = context.user_data.get("division")
        file_path = f"/content/drive/MyDrive/dg/{year}/{tournament}/{division}/holes_avg.png"
        caption = f"📊 Загальний огляд кошиків для дивізіону {division} — турнір {tournament} {year}"

        keyboard = [
            [InlineKeyboardButton("🏅 Кращі гравці на кошиках", callback_data="baskets:top")],
            [InlineKeyboardButton("🔥 Складність кошиків", callback_data="baskets:hardest")],
            [InlineKeyboardButton("📊 Загальний огляд кошиків", callback_data="baskets:overview")],
            [InlineKeyboardButton(f"🔙 Назад до статистики {division}", callback_data="menu:division_stat")],
            [InlineKeyboardButton("🏠 Головне меню", callback_data="back:main")]
        ]

        await send_image_with_caption(query, context, file_path, caption, followup_keyboard=keyboard)


    elif data == "baskets:comparison":
        year = context.user_data.get("year")
        tournament = context.user_data.get("tournament")
        division = context.user_data.get("division")
        file_path = f"/content/drive/MyDrive/dg/{year}/{tournament}/{division}/holes_cmp.png"
        caption = f"⚖️ Порівняння кошиків по дивізіонах — турнір {tournament} {year}"
        await send_image_with_caption(query, context, file_path, caption)


    elif data.startswith("basket:"):
          basket_number = data.split(":")[1]
          chat_id = query.message.chat_id

          # 🧹 Повна очистка: історія + попереднє меню кошиків
          for msg_id in context.user_data.get("history", []):
              try:
                  await context.bot.delete_message(chat_id=chat_id, message_id=msg_id)
              except:
                  pass
          context.user_data["history"] = []

          menu_id = context.user_data.pop("basket_menu_msg_id", None)
          if menu_id:
              try:
                  await context.bot.delete_message(chat_id=chat_id, message_id=menu_id)
              except:
                  pass

          year = context.user_data.get("year")
          tournament = context.user_data.get("tournament")
          file_path = f"/content/drive/MyDrive/dg/{year}/{tournament}/ALL/{basket_number}.png"

          try:
              with open(file_path, "rb") as img:
                  caption = f"📸 Кошик #{basket_number} ({tournament}, {year})"
                  msg = await context.bot.send_photo(chat_id=chat_id, photo=img, caption=caption)
                  context.user_data["history"].append(msg.message_id)

              # 📌 Меню для повторного вибору кошика
              basket_buttons = [InlineKeyboardButton(str(i), callback_data=f"basket:{i}") for i in range(1, 19)]
              keyboard = [basket_buttons[i:i+6] for i in range(0, len(basket_buttons), 6)]
              keyboard.append([
                  InlineKeyboardButton("🔙 Назад", callback_data="division:ALL"),
                  InlineKeyboardButton("🏠 Головне меню", callback_data="back:main")
              ])

              menu_msg = await context.bot.send_message(
                  chat_id=chat_id,
                  text="🔽 Обери інший кошик або повернись:",
                  reply_markup=InlineKeyboardMarkup(keyboard)
              )
              context.user_data["basket_menu_msg_id"] = menu_msg.message_id
              context.user_data["history"].append(menu_msg.message_id)

          except FileNotFoundError:
              msg = await context.bot.send_message(
                  chat_id=chat_id,
                  text=f"❌ Файл `{basket_number}.png` не знайдено у `ALL` для {tournament} {year}.",
                  parse_mode="Markdown"
              )
              context.user_data["history"].append(msg.message_id)
    elif data == "menu:stability":
        chat_id = query.message.chat_id

        # 🧹 Очистка попередніх повідомлень
        for msg_id in context.user_data.get("history", []):
            try:
                await context.bot.delete_message(chat_id=chat_id, message_id=msg_id)
            except:
                pass
        context.user_data["history"] = []

        try:
            await context.bot.delete_message(chat_id=chat_id, message_id=query.message.message_id)
        except:
            pass

        year = context.user_data.get("year")
        tournament = context.user_data.get("tournament")
        division = "ALL"
        file_path = f"/content/drive/MyDrive/dg/{year}/{tournament}/{division}/stability.png"

        try:
            with open(file_path, "rb") as img:
                caption = f"🧱 Стабільність — турнір {tournament} {year}, дивізіон: ALL"
                msg = await context.bot.send_photo(chat_id=chat_id, photo=img, caption=caption)
                context.user_data["history"].append(msg.message_id)

            # 🔁 Меню після фото
            keyboard = [
                [InlineKeyboardButton("🧺 По кошиках", callback_data="menu:baskets")],
                [InlineKeyboardButton("🧱 Стабільність", callback_data="menu:stability")],
                [InlineKeyboardButton("🥏 Ідеальний раунд", callback_data="ideal_round")],
                [InlineKeyboardButton("💩 Жахливий раунд", callback_data="horror_round")],
                [InlineKeyboardButton("🌟 Найкращий раунд", callback_data="best_round")],
                [InlineKeyboardButton("🕳️ Найгірший раунд", callback_data="worst_round")],
                [InlineKeyboardButton(f"🔙 Назад до дивізіонів {tournament} {year}", callback_data=f"back:tournament:{tournament}")],
                [InlineKeyboardButton("🏠 Main Menu", callback_data="back:main")]
            ]
            menu_msg = await context.bot.send_message(
                chat_id=chat_id,
                text="⬇️ Обери наступну опцію:",
                reply_markup=InlineKeyboardMarkup(keyboard)
            )
            context.user_data["history"].append(menu_msg.message_id)

        except FileNotFoundError:
            msg = await context.bot.send_message(
                chat_id=chat_id,
                text=f"❌ Файл `{file_path}` не знайдено.",
                parse_mode="Markdown"
            )
            context.user_data["history"].append(msg.message_id)


    elif data == "ideal_round":
        await send_round_table(
            query, context,
            "ideal_round",
            "📊 Таблиця найкращих результатів на кожному кошику дивізіону {division} — {tournament} {year}"
        )

    elif data == "horror_round":
        await send_round_table(
            query, context,
            "horror_round",
            "📊 Таблиця найгірших результатів на кожному кошику дивізіону {division} — {tournament} {year}"
        )

    elif data == "best_round":
        await send_round_table(
            query, context,
            "best_round",
            "📊 Найкращі раунди учасників дивізіону {division} — {tournament} {year}"
        )

    elif data == "worst_round":
        await send_round_table(
            query, context,
            "worst_round",
            "📊 Найгірші раунди учасників дивізіону {division} — {tournament} {year}"
        )

    elif data == "restart:player_mask":
      context.user_data["awaiting_player_mask"] = True
      await query.edit_message_text("🔍 Введи частину імені гравця для нового пошуку:")


    elif data.startswith("division:"):
        division = data.split(":")[1]
        context.user_data["division"] = division
        year = context.user_data["year"]
        tournament = context.user_data["tournament"]
        path = os.path.join(BASE_DIR, year, tournament, division)
        players = [d for d in os.listdir(path) if os.path.isdir(os.path.join(path, d))]

        selected_player = context.user_data.get("player")

        keyboard = []
        for p in sorted(players):
            if p == selected_player:
                display_name = f"{p}"
            else:
                display_name = p
            keyboard.append([InlineKeyboardButton(display_name, callback_data=f"player:{p}")])
        keyboard.append([InlineKeyboardButton("📦 Статистика дивізіону", callback_data="menu:division_stat")])
        keyboard.append([InlineKeyboardButton(f"🔙 Назад до дивізіонів {tournament} {year}", callback_data=f"back:tournament:{tournament}")])
        keyboard.append([InlineKeyboardButton("🏠 Головне меню", callback_data="back:main")])

        await query.edit_message_text(
            f"🥏 Гравці ({division}) {tournament} {year}:",
            reply_markup=InlineKeyboardMarkup(keyboard)
        )
    elif data == "menu:division_stat":
        year = context.user_data.get("year")
        tournament = context.user_data.get("tournament")
        division = context.user_data.get("division")
        context.user_data.pop("player", None)  # 🔧 очищає контекст гравця

        if not all([year, tournament, division]):
            await query.edit_message_text("⚠️ Немає повної інформації для дивізіону.")
            return

        keyboard = [
            [InlineKeyboardButton("🧺 По кошиках", callback_data="menu:baskets")],
            [InlineKeyboardButton("🥏 Ідеальний раунд", callback_data="ideal_round")],
            [InlineKeyboardButton("💩 Жахливий раунд", callback_data="horror_round")],
            [InlineKeyboardButton("🌟 Найкращий раунд", callback_data="best_round")],
            [InlineKeyboardButton("🕳️ Найгірший раунд", callback_data="worst_round")],
            [InlineKeyboardButton(f"🔙 Назад до гравців ({division})", callback_data=f"division:{division}")],
            [InlineKeyboardButton("🏠 Головне меню", callback_data="back:main")]
        ]


        await query.edit_message_text(
            f"📦 Статистика для дивізіону *{division}* — {tournament} {year}\n\nОберіть опцію:",
            parse_mode="Markdown",
            reply_markup=InlineKeyboardMarkup(keyboard)
        )

    elif data == "back:main":
      chat_id = query.message.chat_id
      history = context.user_data.get("history", [])

      # 🧹 Видаляємо всі повідомлення з історії
      for msg_id in history:
          try:
              await context.bot.delete_message(chat_id=chat_id, message_id=msg_id)
          except:
              pass

      # 🗑️ Видаляємо поточне повідомлення, якщо воно ще є
      try:
          await context.bot.delete_message(chat_id=chat_id, message_id=query.message.message_id)
      except:
          pass

      context.user_data["history"] = []

      # 🔄 Надсилаємо головне меню
      keyboard = [
          [InlineKeyboardButton("📊 По турніру", callback_data="menu:tournament")],
          [InlineKeyboardButton("👤 По гравцю", callback_data="menu:player")],
          [InlineKeyboardButton("🧹 Очистити чат", callback_data="clear_chat")]
      ]

      msg = await context.bot.send_message(
          chat_id=chat_id,
          text="Обери опцію:",
          reply_markup=InlineKeyboardMarkup(keyboard)
      )
      context.user_data["history"].append(msg.message_id)


    elif data == "clear_chat":
        chat_id = query.message.chat_id
        history = context.user_data.get("history", [])

        reply_msg_id = context.user_data.get("reply_keyboard_msg_id")  # ← отримуємо системне повідомлення

        for msg_id in history:
            if msg_id == reply_msg_id:
                continue  # ❌ Пропускаємо, не видаляємо системне
            try:
                await context.bot.delete_message(chat_id=chat_id, message_id=msg_id)
            except:
                pass

        try:
            if query.message.message_id != reply_msg_id:
                await context.bot.delete_message(chat_id=chat_id, message_id=query.message.message_id)
        except:
            pass

        context.user_data["history"] = []

        msg = await context.bot.send_message(chat_id=chat_id, text="✅ Чат очищено ^_^")
        context.user_data["history"].append(msg.message_id)

        await start(update, context)


    elif data == "menu:player":
        context.user_data["awaiting_player_mask"] = True

        all_players = set()
        for year in os.listdir(BASE_DIR):
            year_path = os.path.join(BASE_DIR, year)
            if not os.path.isdir(year_path):
                continue
            for tournament in os.listdir(year_path):
                tournament_path = os.path.join(year_path, tournament)
                if not os.path.isdir(tournament_path):
                    continue
                for division in os.listdir(tournament_path):
                    division_path = os.path.join(tournament_path, division)
                    if not os.path.isdir(division_path):
                        continue
                    for player in os.listdir(division_path):
                        player_path = os.path.join(division_path, player)
                        if os.path.isdir(player_path):
                            all_players.add(player)

        if not all_players:
            await query.edit_message_text("❌ Гравців не знайдено.")
            return


        sorted_players = sorted(all_players)
        keyboard = []
        row = []
        for i, name in enumerate(sorted_players, 1):
            row.append(InlineKeyboardButton(name, callback_data=f"player_lookup:{name}"))
            if i % 2 == 0:
                keyboard.append(row)
                row = []
        if row:
            keyboard.append(row)
        keyboard.append([InlineKeyboardButton("🏠 Головне меню", callback_data="back:main")])

        # ⏺️ Спочатку кнопки
        await query.edit_message_text("Обери гравця зі списку:", reply_markup=InlineKeyboardMarkup(keyboard))

        # 📩 Надсилаємо підказку
        hint_msg = await context.bot.send_message(
            chat_id=query.message.chat_id,
            text="📝 Виберіть гравця зі списку або введіть частину імені у чат 👇"
        )

        context.user_data["player_hint_msg_id"] = hint_msg.message_id  # 💾 Запам’ятовуємо для ручного видалення
        context.user_data.setdefault("history", []).append(hint_msg.message_id)

        # ⏳ Чекаємо 10 секунд і видаляємо
        async def auto_delete_hint():
            await asyncio.sleep(10)
            try:
                await context.bot.delete_message(chat_id=query.message.chat_id, message_id=hint_msg.message_id)
            except:
                pass

        asyncio.create_task(auto_delete_hint())



    elif data.startswith("player_lookup:"):
          # 👉 Сюди вставити блок видалення
        hint_id = context.user_data.pop("player_hint_msg_id", None)
        if hint_id:
            try:
                await context.bot.delete_message(chat_id=query.message.chat_id, message_id=hint_id)
            except:
                pass
        player = data.split(":")[1]
        context.user_data["player"] = player
        tournaments_found = []
        hint_id = context.user_data.pop("player_hint_msg_id", None)
        if hint_id:
            try:
                await context.bot.delete_message(chat_id=query.message.chat_id, message_id=hint_id)
            except Exception as e:
                print("Не вдалося видалити підказку:", e)

        player = data.split(":")[1]
        context.user_data["player"] = player
        tournaments_found = []

        for year in os.listdir(BASE_DIR):
            year_path = os.path.join(BASE_DIR, year)
            if not os.path.isdir(year_path):
                continue
            for tournament in os.listdir(year_path):
                tournament_path = os.path.join(year_path, tournament)
                if not os.path.isdir(tournament_path):
                    continue
                for division in os.listdir(tournament_path):
                    division_path = os.path.join(tournament_path, division)
                    if not os.path.isdir(division_path):
                        continue
                    player_path = os.path.join(division_path, player)
                    if os.path.isdir(player_path):
                        context.user_data["division"] = division
                        tournaments_found.append((year, tournament))
                        break

        if not tournaments_found:
            await query.edit_message_text(f"❌ Гравець {player} не знайдений у жодному турнірі.")
            return

        keyboard = [
            [InlineKeyboardButton(f"{t} ({y})", callback_data=f"tournament_lookup:{y}:{t}")]
            for y, t in sorted(tournaments_found)
        ]

        keyboard.append([InlineKeyboardButton("🔙 Назад до гравців", callback_data="menu:player")])
        keyboard.append([InlineKeyboardButton("🏠 Головне меню", callback_data="back:main")])


        await query.edit_message_text(
            f"📅 Турніри, у яких брав участь {player}:",
            reply_markup=InlineKeyboardMarkup(keyboard)
        )

    elif data.startswith("tournament_lookup:"):
        _, year, tournament = data.split(":")
        context.user_data["year"] = year
        context.user_data["tournament"] = tournament
        player = context.user_data["player"]
        path = os.path.join(BASE_DIR, year, tournament)

        # Знаходимо дивізіон гравця
        divisions = [
            d for d in os.listdir(path)
            if os.path.isdir(os.path.join(path, d, player))
        ]

        if not divisions:
            await query.edit_message_text(
                f"❌ Гравець {player} не знайдений у турнірі {tournament} {year}."
            )
            return

        # Якщо знайшли один — автоматично переходимо
        division = divisions[0]
        context.user_data["division"] = division

        keyboard = [
            [InlineKeyboardButton("🥏 Ідеальний раунд", callback_data="ideal_round")],
            [InlineKeyboardButton("💩 Жахливий раунд", callback_data="horror_round")],
            [InlineKeyboardButton("🌟 Найкращий раунд", callback_data="best_round")],
            [InlineKeyboardButton("🕳️ Найгірший раунд", callback_data="worst_round")],
        ]

        # Додаємо кнопку лише якщо division не дорівнює "ALL"
        if player is not None and division != "ALL":
            keyboard.append([InlineKeyboardButton("📈 Успішність на кошиках", callback_data="player:percent")])

        # Завершальні кнопки
        keyboard.extend([
            [InlineKeyboardButton(f"🔙 Назад до турнірів гравця ({player})", callback_data=f"player_lookup:{player}")],
            [InlineKeyboardButton(f"📋 Перейти до турніру {tournament} {year}", callback_data=f"tournament:{tournament}")],
            [InlineKeyboardButton("🏠 Головне меню", callback_data="back:main")]
        ])


        await query.edit_message_text(
            f"✅ Обрано: {player} ({division}, {tournament} {year})\n\nОберіть опцію:",
            reply_markup=InlineKeyboardMarkup(keyboard)
        )


async def handle_text(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """
    Обробляє текстові повідомлення користувача, коли очікується пошук гравця за маскою.

    :param update: об’єкт Update, що містить повідомлення користувача
    :type update: telegram.Update
    :param context: контекст виконання, що містить user_data та методи для взаємодії з ботом
    :type context: telegram.ext.ContextTypes.DEFAULT_TYPE
    :return: None
    :rtype: None

     Алгоритм:

     - Перевіряє, чи бот очікує введення маски гравця (`awaiting_player_mask`).
     - Зберігає повідомлення користувача та запускає асинхронне видалення цього повідомлення через 5 секунд.
     - Видаляє тимчасову підказку, якщо вона залишилась.
     - Виконує пошук гравців у файловій структурі турнірів:

         - перебирає всі роки, турніри та дивізіони;
         - додає гравців, чиє ім’я містить введену маску.

     - Якщо гравців не знайдено:

         - надсилає повідомлення про помилку;
         - знову активує режим очікування маски.

     - Якщо гравці знайдені:

         - формує клавіатуру з кнопками для кожного гравця (по 2 у рядку);
         - додає кнопки для повторного запиту та повернення в головне меню;
         - надсилає повідомлення зі списком знайдених гравців.

     - Оновлює `context.user_data["history"]`, щоб контролювати видалення повідомлень.
    """
    if context.user_data.get("awaiting_player_mask"):
        mask = update.message.text.strip().lower()

        # ⏳ Затримка 5 секунд, потім видалення повідомлення користувача
        async def auto_delete_user_input(msg_id):
            await asyncio.sleep(5)
            try:
                await context.bot.delete_message(
                    chat_id=update.effective_chat.id,
                    message_id=msg_id
                )
            except:
                pass

        user_msg_id = update.message.message_id
        context.user_data.setdefault("history", []).append(user_msg_id)
        asyncio.create_task(auto_delete_user_input(user_msg_id))

        # 🧽 Видалити тимчасову підказку, якщо залишилась
        hint_id = context.user_data.pop("player_hint_msg_id", None)
        if hint_id:
            try:
                await context.bot.delete_message(chat_id=update.effective_chat.id, message_id=hint_id)
            except:
                pass

        # ⏱️ Далі йде звична логіка пошуку
        all_players = set()
        for year in os.listdir(BASE_DIR):
            for tournament in os.listdir(os.path.join(BASE_DIR, year)):
                for division in os.listdir(os.path.join(BASE_DIR, year, tournament)):
                    for player in os.listdir(os.path.join(BASE_DIR, year, tournament, division)):
                        if mask in player.lower():
                            all_players.add(player)

        context.user_data["awaiting_player_mask"] = False

        if not all_players:
            msg = await update.message.reply_text("❌ Гравців не знайдено. Спробуй ще раз:")
            context.user_data.setdefault("history", []).append(msg.message_id)
            context.user_data["awaiting_player_mask"] = True
            return

        keyboard = []
        row = []
        for i, name in enumerate(sorted(all_players), 1):
            button = InlineKeyboardButton(name, callback_data=f"player_lookup:{name}")
            row.append(button)
            if i % 2 == 0:
                keyboard.append(row)
                row = []
        if row:
            keyboard.append(row)

        keyboard.append([InlineKeyboardButton("🔁 Спробувати інший запит", callback_data="restart:player_mask")])
        keyboard.append([InlineKeyboardButton("🏠 Головне меню", callback_data="back:main")])

        msg = await update.message.reply_text(
            "🔎 Гравці, що збігаються:",
            reply_markup=InlineKeyboardMarkup(keyboard)
        )
        context.user_data.setdefault("history", []).append(msg.message_id)

# Sphinx quickstart should be run from a shell, not inside a module.
# Example PowerShell command to create docs (run in project root):
#   sphinx-quickstart docs -q -p "BotProject" -a "Anton" --ext-autodoc --sep

"""
Модуль запуску Telegram‑бота.

Функціонал:
- створює застосунок на основі токена;
- реєструє обробник команди /start;
- реєструє центральний обробник callback‑кнопок;
- запускає цикл опитування (polling).
"""
app = ApplicationBuilder().token(TOKEN).build()
app.add_handler(CommandHandler("start", start))
app.add_handler(CallbackQueryHandler(button_handler))
app.run_polling()