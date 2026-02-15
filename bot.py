import asyncio
import logging
import os
from datetime import datetime

import pandas as pd
from aiogram import Bot, Dispatcher, F
from aiogram.types import (
    Message,
    CallbackQuery,
    InlineKeyboardMarkup,
    InlineKeyboardButton
)
from aiogram.filters import Command

# =====================
# НАСТРОЙКИ
# =====================

TOKEN = os.getenv("BOT_TOKEN")

if not TOKEN:
    raise ValueError("BOT_TOKEN не найден! Проверь переменные в Railway.")

bot = Bot(token=TOKEN)
dp = Dispatcher()

EXCEL_FILE = "requests.xlsx"

# Telegram ID снабженцев
SUPPLY_USERS = {
    "Никита П.": 111111111,
    "Дмитрий М.": 222222222,
    "Николай К.": 333333333,
}

# Хранилище временных заявок
user_requests = {}

# =====================
# ИНИЦИАЛИЗАЦИЯ EXCEL
# =====================

def init_excel():
    if not os.path.exists(EXCEL_FILE):
        df = pd.DataFrame(columns=[
            "ID",
            "Дата",
            "Начальник",
            "ID_начальника",
            "Текст",
            "Ответственный",
            "Статус",
            "Дата_статуса"
        ])
        df.to_excel(EXCEL_FILE, index=False)

init_excel()

# =====================
# КОМАНДА /start
# =====================

@dp.message(Command("start"))
async def start(message: Message):
    await message.answer("Напишите текст заявки.")

# =====================
# ПРИЁМ ЗАЯВКИ
# =====================

@dp.message()
async def new_request(message: Message):
    text = message.text

    user_requests[message.from_user.id] = {
        "text": text,
        "chief_name": message.from_user.full_name,
        "chief_id": message.from_user.id
    }

    keyboard = InlineKeyboardMarkup(
        inline_keyboard=[
            [
                InlineKeyboardButton(
                    text=name,
                    callback_data=f"assign_{name}"
                )
            ] for name in SUPPLY_USERS.keys()
        ]
    )

    await message.answer("Выберите снабженца:", reply_markup=keyboard)

# =====================
# НАЗНАЧЕНИЕ СНАБЖЕНЦА
# =====================

@dp.callback_query(F.data.startswith("assign_"))
async def assign_supply(callback: CallbackQuery):
    name = callback.data.split("_")[1]
    supply_id = SUPPLY_USERS[name]

    data = user_requests.get(callback.from_user.id)

    if not data:
        await callback.answer("Заявка не найдена. Отправьте заново.")
        return

    df = pd.read_excel(EXCEL_FILE)

    request_id = len(df) + 1
    now = datetime.now().strftime("%Y-%m-%d %H:%M")

    new_row = {
        "ID": request_id,
        "Дата": now,
        "Начальник": data["chief_name"],
        "ID_начальника": data["chief_id"],
        "Текст": data["text"],
        "Ответственный": name,
        "Статус": "Новая",
        "Дата_статуса": now
    }

    df = pd.concat([df, pd.DataFrame([new_row])], ignore_index=True)
    df.to_excel(EXCEL_FILE, index=False)

    keyboard = InlineKeyboardMarkup(inline_keyboard=[
        [
            InlineKeyboardButton(
                text="В работу",
                callback_data=f"work_{request_id}"
            ),
            InlineKeyboardButton(
                text="Закуплено",
                callback_data=f"done_{request_id}"
            )
        ]
    ])

    await bot.send_message(
        supply_id,
        f"Новая заявка №{request_id}\n\n{data['text']}",
        reply_markup=keyboard
    )

    await callback.message.answer(f"Заявка №{request_id} отправлена {name}.")
    await callback.answer()

    del user_requests[callback.from_user.id]

# =====================
# СТАТУС "В РАБОТЕ"
# =====================

@dp.callback_query(F.data.startswith("work_"))
async def set_in_work(callback: CallbackQuery):
    request_id = int(callback.data.split("_")[1])

    df = pd.read_excel(EXCEL_FILE)

    df.loc[df["ID"] == request_id, "Статус"] = "В работе"
    df.loc[df["ID"] == request_id, "Дата_статуса"] = datetime.now().strftime("%Y-%m-%d %H:%M")
    df.to_excel(EXCEL_FILE, index=False)

    chief_id = int(df[df["ID"] == request_id]["ID_начальника"].values[0])

    await bot.send_message(chief_id, f"Заявка №{request_id} взята в работу.")
    await callback.answer("Статус обновлён")

# =====================
# СТАТУС "ЗАКУПЛЕНО"
# =====================

@dp.callback_query(F.data.startswith("done_"))
async def set_done(callback: CallbackQuery):
    request_id = int(callback.data.split("_")[1])

    df = pd.read_excel(EXCEL_FILE)

    df.loc[df["ID"] == request_id, "Статус"] = "Закуплено"
    df.loc[df["ID"] == request_id, "Дата_статуса"] = datetime.now().strftime("%Y-%m-%d %H:%M")
    df.to_excel(EXCEL_FILE, index=False)

    chief_id = int(df[df["ID"] == request_id]["ID_начальника"].values[0])

    await bot.send_message(chief_id, f"Заявка №{request_id} выполнена ✅")
    await callback.answer("Статус обновлён")

# =====================
# ЗАПУСК
# =====================

async def main():
    logging.basicConfig(level=logging.INFO)
    await dp.start_polling(bot)

if __name__ == "__main__":
    asyncio.run(main())
