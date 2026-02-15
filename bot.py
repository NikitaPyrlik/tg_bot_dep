import os
import logging
from datetime import datetime
from aiogram import Bot, Dispatcher, executor, types
from aiogram.types import InlineKeyboardMarkup, InlineKeyboardButton
from openpyxl import Workbook, load_workbook

# ================== НАСТРОЙКИ ==================
BOT_TOKEN = os.getenv("BOT_TOKEN")

REQUESTS_FILE = "requests.xlsx"
SUPPLY_FILE = "supply_users.xlsx"
FILES_DIR = "request_files"

STATUSES = {
    "принята": "🟡 Принята",
    "в_работе": "🔵 В работе",
    "закуплено": "🟢 Закуплено",
}
# ===============================================

logging.basicConfig(level=logging.INFO)

bot = Bot(token=BOT_TOKEN)
dp = Dispatcher(bot)

os.makedirs(FILES_DIR, exist_ok=True)

# ---------- Excel: заявки ----------
def init_requests():
    if not os.path.exists(REQUESTS_FILE):
        wb = Workbook()
        ws = wb.active
        ws.title = "Заявки"
        ws.append([
            "№",
            "Дата создания",
            "Автор",
            "Автор ID",
            "Комментарий",
            "Файл",
            "Статус",
            "Ответственный",
            "Дата статуса",
        ])
        wb.save(REQUESTS_FILE)


def add_request(row):
    wb = load_workbook(REQUESTS_FILE)
    ws = wb.active
    ws.append(row)
    wb.save(REQUESTS_FILE)
    return ws.max_row - 1


# ---------- Excel: снабженцы ----------
def init_supply():
    if not os.path.exists(SUPPLY_FILE):
        wb = Workbook()
        ws = wb.active
        ws.title = "Снабжение"
        ws.append(["Telegram ID", "Имя"])
        wb.save(SUPPLY_FILE)


def add_supply_user(user):
    wb = load_workbook(SUPPLY_FILE)
    ws = wb.active
    for row in ws.iter_rows(min_row=2, values_only=True):
        if row[0] == user.id:
            return
    ws.append([user.id, user.full_name])
    wb.save(SUPPLY_FILE)


def get_supply_users():
    wb = load_workbook(SUPPLY_FILE)
    ws = wb.active
    return [
        {"id": row[0], "name": row[1]}
        for row in ws.iter_rows(min_row=2, values_only=True)
    ]


init_requests()
init_supply()

# ---------- Кнопки ----------
def supply_keyboard(author_id):
    kb = InlineKeyboardMarkup(row_width=1)
    for user in get_supply_users():
        kb.add(
            InlineKeyboardButton(
                text=f"👷 {user['name']}",
                callback_data=f"assign:{author_id}:{user['id']}"
            )
        )
    return kb


def status_keyboard(request_id):
    kb = InlineKeyboardMarkup(row_width=2)
    kb.add(
        InlineKeyboardButton(
            "🔵 В работу",
            callback_data=f"status:{request_id}:в_работе"
        ),
        InlineKeyboardButton(
            "🟢 Закуплено",
            callback_data=f"status:{request_id}:закуплено"
        ),
    )
    return kb


# ---------- Хэндлеры ----------
@dp.message_handler(commands=["start"])
async def start(message: types.Message):
    add_supply_user(message.from_user)
    await message.answer(
        "👋 Бот заявок снабжения\n\n"
        "📎 Отправьте файл заявки\n"
        "✏️ Комментарий — в подписи к файлу"
    )


@dp.message_handler(content_types=types.ContentType.DOCUMENT)
async def handle_document(message: types.Message):
    doc = message.document
    comment = message.caption or "Без комментария"

    file = await bot.get_file(doc.file_id)
    filename = f"{datetime.now():%Y%m%d_%H%M%S}_{doc.file_name}"
    path = os.path.join(FILES_DIR, filename)
    await bot.download_file(file.file_path, path)

    dp.current_state(user=message.from_user.id).update_data(
        pending={
            "author": message.from_user,
            "comment": comment,
            "document_id": doc.file_id,
            "filename": filename,
        }
    )

    await message.answer(
        "👷 Кому отправить заявку?",
        reply_markup=supply_keyboard(message.from_user.id)
    )


@dp.callback_query_handler(lambda c: c.data.startswith("assign:"))
async def assign_request(callback: types.CallbackQuery):
    _, author_id, supply_id = callback.data.split(":")
    supply_id = int(supply_id)

    state = dp.current_state(user=int(author_id))
    data = (await state.get_data()).get("pending")

    if not data:
        await callback.answer("Заявка не найдена", show_alert=True)
        return

    supply_user = next(
        u for u in get_supply_users() if u["id"] == supply_id
    )

    now = datetime.now().strftime("%d.%m.%Y %H:%M")

    request_id = add_request([
        "",
        now,
        data["author"].full_name,
        data["author"].id,
        data["comment"],
        data["filename"],
        STATUSES["принята"],
        supply_user["name"],
        now,
    ])

    wb = load_workbook(REQUESTS_FILE)
    ws = wb.active
    ws.cell(row=request_id + 1, column=1).value = request_id
    wb.save(REQUESTS_FILE)

    await bot.send_document(
        supply_user["id"],
        data["document_id"],
        caption=(
            f"📦 *Заявка №{request_id}*\n\n"
            f"👤 {data['author'].full_name}\n"
            f"📝 {data['comment']}\n"
            f"📊 {STATUSES['принята']}"
        ),
        parse_mode="Markdown",
        reply_markup=status_keyboard(request_id)
    )

    await callback.message.edit_text(
        f"✅ Заявка отправлена: {supply_user['name']}"
    )
    await callback.answer()


@dp.callback_query_handler(lambda c: c.data.startswith("status:"))
async def change_status(callback: types.CallbackQuery):
    _, request_id, status_key = callback.data.split(":")
    request_id = int(request_id)
    new_status = STATUSES[status_key]
    now = datetime.now().strftime("%d.%m.%Y %H:%M")

    wb = load_workbook(REQUESTS_FILE)
    ws = wb.active

    for row in ws.iter_rows(min_row=2):
        if row[0].value == request_id:
            row[6].value = new_status
            row[8].value = now
            author_id = row[3].value
            responsible = row[7].value
            wb.save(REQUESTS_FILE)

            await bot.send_message(
                author_id,
                (
                    f"📦 *Заявка №{request_id}*\n\n"
                    f"📊 Статус: {new_status}\n"
                    f"👷 Ответственный: {responsible}\n"
                    f"🕒 Дата: {now}"
                ),
                parse_mode="Markdown"
            )
            break

    await callback.message.edit_caption(
        callback.message.caption + f"\n\n📊 {new_status}",
        parse_mode="Markdown"
    )
    await callback.answer("Статус обновлён ✅")


if __name__ == "__main__":
    executor.start_polling(dp, skip_updates=True)
