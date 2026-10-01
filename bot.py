"""Лид-Радар: Telegram-бот для поиска клиентов на услуги по отзывам (Яндекс.Карты).

Пользователям: кнопочный поиск, подписка, профиль.
Администратору: статистика, платежи, пользователи, рассылка, настройки, экспорт.

Настройки читаются из переменных окружения или из файла .env рядом с ботом:
    BOT_TOKEN=...            токен от @BotFather (обязательно)
    ADMIN_IDS=123,456        Telegram ID администраторов (узнать: команда /id в боте)
    PAY_DETAILS=...          реквизиты для оплаты, показываются покупателю
    BRAND=Лид-Радар          название продукта (необязательно)
"""
import asyncio
import html
import os
import sqlite3
import tempfile
import time
from datetime import datetime

from aiogram import BaseMiddleware, Bot, Dispatcher, F
from aiogram.client.default import DefaultBotProperties
from aiogram.exceptions import TelegramBadRequest
from aiogram.filters import Command
from aiogram.types import (BotCommand, CallbackQuery, FSInputFile, InlineKeyboardButton,
                           InlineKeyboardMarkup, Message)
from aiogram.utils.keyboard import InlineKeyboardBuilder
from openpyxl import Workbook

from cli import badge, body, build_xlsx, header
from parser import search_async


def load_env(path=".env"):
    if os.path.exists(path):
        for line in open(path, encoding="utf-8"):
            line = line.strip()
            if line and not line.startswith("#") and "=" in line:
                k, v = line.split("=", 1)
                os.environ.setdefault(k.strip(), v.strip().strip('"').strip("'"))


load_env()
TOKEN = os.environ["BOT_TOKEN"]
ADMINS = {int(x) for x in os.getenv("ADMIN_IDS", "").split(",") if x.strip()}
PAY_DETAILS = os.getenv("PAY_DETAILS", "").strip()
BRAND = os.getenv("BRAND", "Лид-Радар")
DB_PATH = os.getenv("DB_PATH", "data.db")

# ───────────────────────── база данных ─────────────────────────
DB = sqlite3.connect(DB_PATH, check_same_thread=False)
DB.row_factory = sqlite3.Row
DB.executescript("""
CREATE TABLE IF NOT EXISTS users(id INTEGER PRIMARY KEY, username TEXT, name TEXT, joined INTEGER,
    sub_until INTEGER DEFAULT 0, trial_used INTEGER DEFAULT 0, banned INTEGER DEFAULT 0, searches INTEGER DEFAULT 0);
CREATE TABLE IF NOT EXISTS payments(id INTEGER PRIMARY KEY AUTOINCREMENT, user_id INTEGER, amount INTEGER,
    days INTEGER, status TEXT, method TEXT, created INTEGER, confirmed INTEGER);
CREATE TABLE IF NOT EXISTS searches(id INTEGER PRIMARY KEY AUTOINCREMENT, user_id INTEGER, query TEXT,
    found INTEGER, leads INTEGER, created INTEGER);
CREATE TABLE IF NOT EXISTS settings(key TEXT PRIMARY KEY, value TEXT);
""")
DB.commit()

DAY = 86400
STATUS = {"pending": "⏳ создан", "review": "🕓 на проверке", "paid": "✅ оплачен", "rejected": "❌ отклонён"}


def now() -> int:
    return int(time.time())


def setting(key: str, default: int) -> int:
    r = DB.execute("SELECT value FROM settings WHERE key=?", (key,)).fetchone()
    return int(r["value"]) if r else default


def set_setting(key: str, value: int):
    DB.execute("INSERT INTO settings(key,value) VALUES(?,?) ON CONFLICT(key) DO UPDATE SET value=excluded.value",
               (key, str(value)))
    DB.commit()


def price() -> int:
    return setting("price", 30)


def period() -> int:
    return setting("period", 30)


def trial_on() -> bool:
    return bool(setting("trial", 1))


def touch(u):
    DB.execute("INSERT OR IGNORE INTO users(id,username,name,joined) VALUES(?,?,?,?)",
               (u.id, u.username or "", u.full_name or "", now()))
    DB.execute("UPDATE users SET username=?, name=? WHERE id=?", (u.username or "", u.full_name or "", u.id))
    DB.commit()


def get_user(uid: int):
    return DB.execute("SELECT * FROM users WHERE id=?", (uid,)).fetchone()


def active(u) -> bool:
    return bool(u) and u["sub_until"] > now()


def grant_days(uid: int, days: int) -> int:
    u = get_user(uid)
    base = max(now(), u["sub_until"])
    until = base + days * DAY
    DB.execute("UPDATE users SET sub_until=? WHERE id=?", (until, uid))
    DB.commit()
    return until


def can_search(uid: int):
    """(можно ли искать, это пробный поиск?)"""
    u = get_user(uid)
    if uid in ADMINS or active(u):
        return True, False
    if trial_on() and not u["trial_used"]:
        return True, True
    return False, False


def create_payment(uid: int) -> int:
    r = DB.execute("SELECT id FROM payments WHERE user_id=? AND status='pending' ORDER BY id DESC LIMIT 1",
                   (uid,)).fetchone()
    if r:
        DB.execute("UPDATE payments SET amount=?, days=? WHERE id=?", (price(), period(), r["id"]))
        DB.commit()
        return r["id"]
    cur = DB.execute("INSERT INTO payments(user_id,amount,days,status,method,created) VALUES(?,?,?,?,?,?)",
                     (uid, price(), period(), "pending", "manual", now()))
    DB.commit()
    return cur.lastrowid


def settle(pid: int, ok: bool):
    """Подтверждает или отклоняет оплату. Сюда же можно подключить платёжную систему."""
    p = DB.execute("SELECT * FROM payments WHERE id=?", (pid,)).fetchone()
    if not p or p["status"] not in ("pending", "review"):
        return None
    DB.execute("UPDATE payments SET status=?, confirmed=? WHERE id=?",
               ("paid" if ok else "rejected", now(), pid))
    DB.commit()
    if ok:
        grant_days(p["user_id"], p["days"])
    return p


def dt(ts, with_time=False) -> str:
    if not ts:
        return "—"
    return datetime.fromtimestamp(ts).strftime("%d.%m.%Y %H:%M" if with_time else "%d.%m.%Y")


def who(r) -> str:
    return html.escape(f"@{r['username']}" if r["username"] else (r["name"] or str(r["user_id"] if "user_id" in r.keys() else r["id"])))


def one(sql, *args) -> int:
    return DB.execute(sql, args).fetchone()[0] or 0


# ───────────────────────── состояние поиска ─────────────────────────
CATS = ["Барбершоп", "Автомойка", "Шиномонтаж", "Салон красоты", "Маникюр",
        "Стоматология", "Кафе", "Фитнес-клуб", "Ремонт квартир", "Клининг"]
CITIES = ["Москва", "Санкт-Петербург", "Уфа", "Казань", "Екатеринбург",
          "Новосибирск", "Краснодар", "Сочи", "Нижний Новгород", "Самара"]
RATES = [3.5, 4.0, 4.3, 4.5]
CARDS = [20, 50, 100, 200, 500, 0]   # 0 = без ограничений
MINS = [0, 3, 5, 10, 20]
UNLIMITED = 10 ** 6
TRIAL_CARDS = 20

bot = Bot(TOKEN, default=DefaultBotProperties(parse_mode="HTML"))
dp = Dispatcher()
lock = asyncio.Lock()       # один поиск за раз: браузер ест много памяти
SETTINGS: dict[int, dict] = {}
WAIT: dict[int, str] = {}
CAST: dict[int, str] = {}
RUNNING: set[int] = set()

HINTS = {
    "cat": "Напишите сферу. Можно несколько через запятую.",
    "city": "Напишите название города.",
    "rate": "Напишите максимальный рейтинг, например 3.8",
    "cards": "Сколько карточек просмотреть? Напишите любое число, например 250. Ноль значит без ограничений.",
    "min": "Напишите минимум оценок, например 15. Ноль значит любое количество.",
    "a_cast": "Напишите текст рассылки. Его получат все пользователи.",
    "a_price": "Напишите новую цену подписки в рублях, числом.",
    "a_period": "Напишите срок подписки в днях, числом.",
}


def st(uid: int) -> dict:
    return SETTINGS.setdefault(uid, {"cats": [], "city": "", "rating": 4.3, "cards": 50, "min": 5})


def fmt_cards(v: int) -> str:
    return "без лимита" if v == 0 else str(v)


def short_cats(cats: list) -> str:
    first = cats[0] if len(cats) and len(cats[0]) <= 18 else cats[0][:17] + "…"
    return first if len(cats) == 1 else f"{first} +{len(cats) - 1}"


def kb_of(builder: InlineKeyboardBuilder) -> InlineKeyboardMarkup:
    return builder.as_markup()


def btn(text, data) -> InlineKeyboardButton:
    return InlineKeyboardButton(text=text, callback_data=data)


def back_kb(data="home", text="← Меню") -> InlineKeyboardMarkup:
    return InlineKeyboardMarkup(inline_keyboard=[[btn(text, data)]])


async def edit(cb: CallbackQuery, text: str, kb: InlineKeyboardMarkup | None = None):
    try:
        await cb.message.edit_text(text, reply_markup=kb)
    except TelegramBadRequest:
        pass  # текст не изменился


# ───────────────────────── тексты для покупателя ─────────────────────────
def pitch() -> str:
    trial = (f"\n\n🎁 <b>Первый поиск бесплатно</b> (до {TRIAL_CARDS} карточек): убедитесь сами."
             if trial_on() else "")
    return (
        f"💎 <b>{html.escape(BRAND)}</b>\n"
        "Клиенты на услуги по отзывам: список за минуты, а не за вечер.\n\n"
        "Вы выбираете сферу и город. Бот сам просматривает карточки на Яндекс.Картах, отбирает компании "
        "с низким рейтингом и присылает готовую таблицу:\n\n"
        "🔥 <b>Приоритет каждого лида</b>: понятно, кому писать первым\n"
        "📞 <b>Телефон, адрес и ссылка</b> на карточку\n"
        "✉️ <b>Готовое первое сообщение</b> под каждую компанию\n"
        "📊 <b>Сводка</b> с ключевыми цифрами\n\n"
        "Никаких ручных проверок и копирования. Всё управление кнопками, работает прямо в Telegram."
        f"{trial}\n\n"
        f"💳 Подписка: <b>{price()} ₽</b> на {period()} дней, любое число поисков."
    )


def home_text(uid: int) -> str:
    u = get_user(uid)
    if uid in ADMINS or active(u):
        left = "без ограничений (администратор)" if uid in ADMINS else f"до {dt(u['sub_until'])}"
        return f"💎 <b>{html.escape(BRAND)}</b>\n\nПодписка активна: {left}.\nВыберите действие."
    return pitch()


def home_kb(uid: int) -> InlineKeyboardMarkup:
    u = get_user(uid)
    paid = uid in ADMINS or active(u)
    trial = trial_on() and not u["trial_used"]
    b = InlineKeyboardBuilder()
    b.button(text="🎁 Попробовать бесплатно" if (trial and not paid) else "🔎 Найти клиентов", callback_data="search")
    b.button(text="💎 Подписка", callback_data="sub")
    b.button(text="👤 Профиль", callback_data="prof")
    b.button(text="ℹ️ Как это работает", callback_data="how")
    if uid in ADMINS:
        b.button(text="🛠 Админ-панель", callback_data="a:home")
    b.adjust(1, 2, 1, 1)
    return kb_of(b)


HOW = (
    "ℹ️ <b>Как это работает</b>\n\n"
    "1️⃣ Выбираете сферу и город, например «Барбершоп, Уфа»\n"
    "2️⃣ Задаёте порог рейтинга и сколько карточек просмотреть\n"
    "3️⃣ Бот собирает компании, которым нужна работа с отзывами\n"
    "4️⃣ Присылает Excel: сводка, лиды по приоритету и готовые сообщения\n\n"
    "Поиск занимает несколько минут. Пока он идёт, бот показывает прогресс.\n\n"
    "💡 Сообщения в таблице рассчитаны на честные методы: настоящие отзывы от реальных клиентов."
)


def sub_text(uid: int) -> str:
    u = get_user(uid)
    status = ("администратор" if uid in ADMINS else
              f"активна до {dt(u['sub_until'])}" if active(u) else "не активна")
    return (
        f"💎 <b>Подписка {html.escape(BRAND)}</b>\n\n"
        f"Стоимость: <b>{price()} ₽</b> на {period()} дней\n"
        "В подписку входит:\n"
        "• любое количество поисков\n"
        "• Excel со сводкой, приоритетами и готовыми сообщениями\n"
        "• любые сферы и города\n"
        "• обновления бота\n\n"
        f"Статус: {status}"
    )


def profile_text(uid: int) -> str:
    u = get_user(uid)
    sub = ("администратор" if uid in ADMINS else
           f"активна до {dt(u['sub_until'])}" if active(u) else "не активна")
    return (f"👤 <b>Профиль</b>\n\nID: <code>{uid}</code>\nПодписка: {sub}\n"
            f"Поисков выполнено: {u['searches']}")


# ───────────────────────── панель поиска ─────────────────────────
def summary(uid: int) -> str:
    s = st(uid)
    cats = html.escape(", ".join(s["cats"])) or "— не выбрана"
    return (
        "🔎 <b>Поиск клиентов с низким рейтингом</b>\n"
        "━━━━━━━━━━━━━━━━━━\n"
        f"🏷 Сфера:  {cats}\n"
        f"📍 Город:  {html.escape(s['city']) or '— не выбран'}\n"
        f"⭐ Рейтинг не выше:  {s['rating']}\n"
        f"📄 Смотреть карточек:  {fmt_cards(s['cards'])}\n"
        f"💬 Оценок не меньше:  {s['min']}\n"
        "━━━━━━━━━━━━━━━━━━\n"
        "Выберите параметры кнопками и нажмите «Запустить»."
    )


def main_kb(uid: int) -> InlineKeyboardMarkup:
    s = st(uid)
    b = InlineKeyboardBuilder()
    b.button(text=f"✅ {short_cats(s['cats'])}" if s["cats"] else "🏷 Выбрать сферу", callback_data="m:cat")
    b.button(text=f"✅ {s['city']}" if s["city"] else "📍 Выбрать город", callback_data="m:city")
    b.button(text=f"⭐ Рейтинг до {s['rating']}", callback_data="m:rate")
    b.button(text=f"📄 Карточек: {fmt_cards(s['cards'])}", callback_data="m:cards")
    b.button(text=f"💬 Оценок от {s['min']}", callback_data="m:min")
    b.button(text="🚀 Запустить поиск", callback_data="run")
    b.button(text="🏠 Меню", callback_data="home")
    b.adjust(2, 2, 1, 1, 1)
    return kb_of(b)


def options_kb(options, prefix, selected, custom=None, done_text="← Назад", cols=2, fmt=str):
    b = InlineKeyboardBuilder()
    for i, o in enumerate(options):
        b.button(text=f"{'✅ ' if o in selected else ''}{fmt(o)}", callback_data=f"{prefix}:{i}")
    b.adjust(cols)
    if custom:
        b.row(btn(*custom))
    b.row(btn(done_text, "search"))
    return kb_of(b)


def submenu(uid: int, kind: str):
    s = st(uid)
    if kind == "cat":
        return ("Выберите сферы (можно несколько) и нажмите «Готово».",
                options_kb(CATS, "tc", s["cats"], ("✍️ Своя сфера", "w:cat"), "✅ Готово"))
    if kind == "city":
        return ("Выберите город.", options_kb(CITIES, "sc", [s["city"]], ("✍️ Другой город", "w:city")))
    if kind == "rate":
        return ("Показывать компании с рейтингом не выше:",
                options_kb(RATES, "sr", [s["rating"]], ("✍️ Своё значение", "w:rate"), cols=4))
    if kind == "cards":
        return ("Сколько карточек просмотреть на каждый запрос:",
                options_kb(CARDS, "sn", [s["cards"]], ("✍️ Своё число", "w:cards"), cols=3,
                           fmt=lambda v: "∞ без лимита" if v == 0 else str(v)))
    return ("Минимум оценок у компании:",
            options_kb(MINS, "sm", [s["min"]], ("✍️ Своё число", "w:min"), cols=5))


def apply_text(uid: int, kind: str, text: str) -> bool:
    s = st(uid)
    text = text.strip()
    try:
        if kind == "cat":
            for c in text.split(","):
                c = c.strip()
                if c and c not in s["cats"]:
                    s["cats"].append(c)
        elif kind == "city":
            s["city"] = text
        elif kind == "rate":
            v = float(text.replace(",", "."))
            if not 1 <= v <= 5:
                return False
            s["rating"] = v
        elif kind in ("cards", "min"):
            v = int(text)
            if v < 0:
                return False
            s["cards" if kind == "cards" else "min"] = v
    except ValueError:
        return False
    return True


# ───────────────────────── доступ и общие команды ─────────────────────────
class Access(BaseMiddleware):
    async def __call__(self, handler, event, data):
        u = data.get("event_from_user")
        if u:
            touch(u)
            row = get_user(u.id)
            if row["banned"] and u.id not in ADMINS:
                if isinstance(event, CallbackQuery):
                    return await event.answer("Доступ ограничен.", show_alert=True)
                return await event.answer("Доступ ограничен.")
        return await handler(event, data)


dp.message.outer_middleware(Access())
dp.callback_query.outer_middleware(Access())


@dp.message(Command("id"))
async def cmd_id(m: Message):
    await m.answer(f"Ваш Telegram ID: <code>{m.from_user.id}</code>")


@dp.message(Command("start", "menu"))
async def cmd_start(m: Message):
    WAIT.pop(m.from_user.id, None)
    await m.answer(home_text(m.from_user.id), reply_markup=home_kb(m.from_user.id))


@dp.message(Command("admin"))
async def cmd_admin(m: Message):
    if m.from_user.id in ADMINS:
        await m.answer(admin_home_text(), reply_markup=admin_home_kb())


@dp.callback_query(F.data == "home")
async def cb_home(cb: CallbackQuery):
    WAIT.pop(cb.from_user.id, None)
    await cb.answer()
    await edit(cb, home_text(cb.from_user.id), home_kb(cb.from_user.id))


@dp.callback_query(F.data == "how")
async def cb_how(cb: CallbackQuery):
    await cb.answer()
    await edit(cb, HOW, back_kb())


@dp.callback_query(F.data == "prof")
async def cb_prof(cb: CallbackQuery):
    await cb.answer()
    await edit(cb, profile_text(cb.from_user.id), back_kb())


@dp.callback_query(F.data == "sub")
async def cb_sub(cb: CallbackQuery):
    await cb.answer()
    uid = cb.from_user.id
    u = get_user(uid)
    b = InlineKeyboardBuilder()
    label = "➕ Продлить" if active(u) else "💳 Оформить подписку"
    b.button(text=f"{label} — {price()} ₽", callback_data="pay")
    b.button(text="← Меню", callback_data="home")
    b.adjust(1)
    await edit(cb, sub_text(uid), kb_of(b))


@dp.callback_query(F.data == "pay")
async def cb_pay(cb: CallbackQuery):
    await cb.answer()
    uid = cb.from_user.id
    if not PAY_DETAILS:
        return await edit(cb, "Оплата пока не настроена. Напишите администратору.", back_kb("sub", "← Назад"))
    pid = create_payment(uid)
    text = (f"💳 <b>Оплата подписки</b>\n\nСумма: <b>{price()} ₽</b> на {period()} дней\n\n"
            f"Реквизиты:\n{html.escape(PAY_DETAILS)}\n\n"
            f"В комментарии к переводу укажите код <code>{pid}</code>.\n"
            "После оплаты нажмите кнопку ниже, и подписка включится после проверки.")
    b = InlineKeyboardBuilder()
    b.button(text="✅ Я оплатил(а)", callback_data=f"paid:{pid}")
    b.button(text="← Назад", callback_data="sub")
    b.adjust(1)
    await edit(cb, text, kb_of(b))


@dp.callback_query(F.data.startswith("paid:"))
async def cb_paid(cb: CallbackQuery):
    pid = int(cb.data[5:])
    p = DB.execute("SELECT * FROM payments WHERE id=?", (pid,)).fetchone()
    if not p or p["user_id"] != cb.from_user.id:
        return await cb.answer("Платёж не найден.", show_alert=True)
    if p["status"] != "pending":
        return await cb.answer("Этот платёж уже отправлен на проверку.", show_alert=True)
    DB.execute("UPDATE payments SET status='review' WHERE id=?", (pid,))
    DB.commit()
    await cb.answer()
    await edit(cb, "🕓 <b>Оплата на проверке</b>\n\nКак только платёж подтвердится, подписка включится и "
                   "я пришлю уведомление.", back_kb())
    u = get_user(cb.from_user.id)
    b = InlineKeyboardBuilder()
    b.button(text="✅ Подтвердить", callback_data=f"a:ok:{pid}")
    b.button(text="❌ Отклонить", callback_data=f"a:no:{pid}")
    for admin_id in ADMINS:
        try:
            await bot.send_message(admin_id, f"💳 <b>Новая оплата #{pid}</b>\n{who(u)} (<code>{u['id']}</code>)\n"
                                             f"{p['amount']} ₽ / {p['days']} дн.", reply_markup=kb_of(b))
        except Exception:
            pass


# ───────────────────────── поиск ─────────────────────────
@dp.callback_query(F.data == "search")
async def cb_search(cb: CallbackQuery):
    WAIT.pop(cb.from_user.id, None)
    uid = cb.from_user.id
    ok, trial = can_search(uid)
    await cb.answer()
    if not ok:
        b = InlineKeyboardBuilder()
        b.button(text=f"💳 Оформить подписку — {price()} ₽", callback_data="pay")
        b.button(text="← Меню", callback_data="home")
        b.adjust(1)
        return await edit(cb, "🔒 Пробный поиск использован. Чтобы продолжить, оформите подписку.\n\n" + sub_text(uid),
                          kb_of(b))
    note = f"\n\n🎁 Пробный поиск: до {TRIAL_CARDS} карточек." if trial else ""
    await edit(cb, summary(uid) + note, main_kb(uid))


@dp.callback_query(F.data.startswith("m:"))
async def cb_menu(cb: CallbackQuery):
    await cb.answer()
    text, kb = submenu(cb.from_user.id, cb.data[2:])
    await edit(cb, text, kb)


@dp.callback_query(F.data.startswith("tc:"))
async def cb_toggle_cat(cb: CallbackQuery):
    s = st(cb.from_user.id)
    cat = CATS[int(cb.data[3:])]
    s["cats"].remove(cat) if cat in s["cats"] else s["cats"].append(cat)
    await cb.answer()
    text, kb = submenu(cb.from_user.id, "cat")
    await edit(cb, text, kb)


@dp.callback_query(F.data.regexp(r"^(sc|sr|sn|sm):\d+$"))
async def cb_set(cb: CallbackQuery):
    s = st(cb.from_user.id)
    key, idx = cb.data.split(":")
    idx = int(idx)
    if key == "sc":
        s["city"] = CITIES[idx]
    elif key == "sr":
        s["rating"] = RATES[idx]
    elif key == "sn":
        s["cards"] = CARDS[idx]
    else:
        s["min"] = MINS[idx]
    await cb.answer("Готово")
    await edit(cb, summary(cb.from_user.id), main_kb(cb.from_user.id))


@dp.callback_query(F.data.startswith("w:"))
async def cb_ask(cb: CallbackQuery):
    kind = cb.data[2:]
    WAIT[cb.from_user.id] = kind
    await cb.answer()
    await edit(cb, HINTS[kind], back_kb("search", "← Назад"))


@dp.callback_query(F.data == "run")
async def cb_run(cb: CallbackQuery):
    uid = cb.from_user.id
    s = st(uid)
    ok, trial = can_search(uid)
    if not ok:
        return await cb.answer("Нужна подписка.", show_alert=True)
    if not s["cats"] or not s["city"]:
        return await cb.answer("Сначала выберите сферу и город.", show_alert=True)
    if uid in RUNNING:
        return await cb.answer("Ваш поиск уже идёт.", show_alert=True)
    await cb.answer()
    try:
        await cb.message.edit_reply_markup(reply_markup=None)  # защита от повторного нажатия
    except TelegramBadRequest:
        pass
    await run_search(cb.message, uid, dict(s), trial)


async def run_search(m: Message, uid: int, s: dict, trial: bool):
    RUNNING.add(uid)
    try:
        cards = s["cards"] or UNLIMITED
        if trial:
            cards = min(cards, TRIAL_CARDS)
        if lock.locked():
            await m.answer("Сейчас идёт другой поиск, ваш запрос в очереди.")
        async with lock:
            status = await m.answer("⏳ Запускаю поиск…")
            last = {"line": "Запускаю поиск…"}

            async def updater():
                shown = ""
                while True:
                    await asyncio.sleep(6)
                    if last["line"] != shown:
                        shown = last["line"]
                        try:
                            await status.edit_text(f"⏳ {html.escape(shown)}")
                        except Exception:
                            pass

            task = asyncio.create_task(updater())
            rows = []
            try:
                for cat in s["cats"]:
                    rows += await search_async(cat, s["city"], max_results=cards, max_rating=s["rating"],
                                               min_reviews=s["min"], log=lambda x: last.update(line=x))
            except Exception as e:
                await m.answer(f"Ошибка при поиске: {html.escape(str(e))}")
                return
            finally:
                task.cancel()

            path = os.path.join(tempfile.gettempdir(), f"leads_{uid}.xlsx")
            leads = build_xlsx(rows, path)
            await status.edit_text(f"✅ Готово. Просмотрено: {len(rows)}, подходящих лидов: {len(leads)}.")
            if not rows:
                await m.answer("Ничего не найдено. Возможно, Яндекс не отдал страницу. Попробуйте позже.")
            else:
                DB.execute("INSERT INTO searches(user_id,query,found,leads,created) VALUES(?,?,?,?,?)",
                           (uid, f"{', '.join(s['cats'])} / {s['city']}", len(rows), len(leads), now()))
                DB.execute("UPDATE users SET searches=searches+1 WHERE id=?", (uid,))
                if trial:
                    DB.execute("UPDATE users SET trial_used=1 WHERE id=?", (uid,))
                DB.commit()
                await m.answer_document(FSInputFile(path, filename="leads.xlsx"))
                if leads:
                    lines = [html.escape(f"{badge(r['score'])} · {r['name']} · {r['rating']}★ ({r['reviews']}) · "
                                         f"{r['phone'] or 'без телефона'}") for r in leads[:5]]
                    await m.answer("Кому писать первым:\n\n" + "\n".join(lines))
                if trial:
                    await m.answer(f"🎁 Это был пробный поиск. Чтобы искать дальше, оформите подписку: "
                                   f"{price()} ₽ на {period()} дней.", reply_markup=back_kb("sub", "💎 Подписка"))
    finally:
        RUNNING.discard(uid)
    await m.answer(summary(uid), reply_markup=main_kb(uid))


# ───────────────────────── админ-панель ─────────────────────────
def admin_home_text() -> str:
    return (f"🛠 <b>Админ-панель {html.escape(BRAND)}</b>\n\n"
            f"Пользователей: {one('SELECT COUNT(*) FROM users')}\n"
            f"Активных подписок: {one('SELECT COUNT(*) FROM users WHERE sub_until>?', now())}\n"
            f"Оплат на проверке: {one('SELECT COUNT(*) FROM payments WHERE status=?', 'review')}")


def admin_home_kb() -> InlineKeyboardMarkup:
    b = InlineKeyboardBuilder()
    for t, d in [("📊 Статистика", "a:stats"), ("💳 Платежи", "a:pay"), ("👥 Пользователи", "a:users"),
                 ("📣 Рассылка", "a:cast"), ("⚙️ Настройки", "a:set"), ("📥 Экспорт в Excel", "a:export"),
                 ("🏠 В меню", "home")]:
        b.button(text=t, callback_data=d)
    b.adjust(2, 2, 1, 1)
    return kb_of(b)


def stats_text() -> str:
    t, week, month = now(), now() - 7 * DAY, now() - 30 * DAY
    paid = "FROM payments WHERE status='paid'"
    return (
        "📊 <b>Статистика</b>\n\n"
        f"👥 Пользователей: {one('SELECT COUNT(*) FROM users')} "
        f"(за 7 дней +{one('SELECT COUNT(*) FROM users WHERE joined>?', week)})\n"
        f"💎 Активных подписок: {one('SELECT COUNT(*) FROM users WHERE sub_until>?', t)}\n"
        f"🎁 Использовали пробный поиск: {one('SELECT COUNT(*) FROM users WHERE trial_used=1')}\n"
        f"🚫 Заблокировано: {one('SELECT COUNT(*) FROM users WHERE banned=1')}\n\n"
        f"💰 Выручка всего: {one('SELECT SUM(amount) ' + paid)} ₽ ({one('SELECT COUNT(*) ' + paid)} оплат)\n"
        f"💰 За 30 дней: {one('SELECT SUM(amount) ' + paid + ' AND confirmed>?', month)} ₽\n"
        f"🕓 На проверке: {one('SELECT COUNT(*) FROM payments WHERE status=?', 'review')}\n\n"
        f"🔎 Поисков всего: {one('SELECT COUNT(*) FROM searches')} "
        f"(за 7 дней {one('SELECT COUNT(*) FROM searches WHERE created>?', week)})"
    )


async def admin_screen(cb: CallbackQuery, parts: list):
    act = parts[1]
    back = InlineKeyboardMarkup(inline_keyboard=[[btn("← Админка", "a:home")]])
    if act == "home":
        return await edit(cb, admin_home_text(), admin_home_kb())
    if act == "stats":
        return await edit(cb, stats_text(), back)
    if act == "pay":
        rows = DB.execute("SELECT p.*, u.username, u.name FROM payments p LEFT JOIN users u ON u.id=p.user_id "
                          "ORDER BY p.id DESC LIMIT 10").fetchall()
        lines = [f"#{r['id']} · {who(r)} · {r['amount']} ₽ · {STATUS[r['status']]} · {dt(r['created'], True)}"
                 for r in rows] or ["Платежей пока нет."]
        b = InlineKeyboardBuilder()
        for r in rows:
            if r["status"] in ("pending", "review"):
                b.button(text=f"✅ #{r['id']}", callback_data=f"a:ok:{r['id']}")
                b.button(text=f"❌ #{r['id']}", callback_data=f"a:no:{r['id']}")
        b.adjust(4)
        b.row(btn("← Админка", "a:home"))
        return await edit(cb, "💳 <b>Последние платежи</b>\n\n" + "\n".join(lines), kb_of(b))
    if act == "users":
        rows = DB.execute("SELECT * FROM users ORDER BY joined DESC LIMIT 10").fetchall()
        b = InlineKeyboardBuilder()
        for r in rows:
            sub = f"до {dt(r['sub_until'])}" if r["sub_until"] > now() else "нет подписки"
            name = (f"@{r['username']}" if r["username"] else r["name"] or str(r["id"]))[:16]
            b.button(text=f"{name} · {sub}", callback_data=f"a:u:{r['id']}")
        b.adjust(1)
        b.row(btn("← Админка", "a:home"))
        return await edit(cb, "👥 <b>Последние пользователи</b>\nНажмите на пользователя, чтобы открыть карточку.",
                          kb_of(b))
    if act == "u":
        uid = int(parts[2])
        u = get_user(uid)
        pays = one("SELECT COUNT(*) FROM payments WHERE user_id=? AND status='paid'", uid)
        text = (f"👤 <b>{who(u)}</b>\nID: <code>{uid}</code>\nРегистрация: {dt(u['joined'])}\n"
                f"Подписка до: {dt(u['sub_until']) if u['sub_until'] > now() else 'нет'}\n"
                f"Оплат: {pays}\nПоисков: {u['searches']}\n"
                f"Пробный поиск: {'использован' if u['trial_used'] else 'не использован'}\n"
                f"Статус: {'🚫 заблокирован' if u['banned'] else 'активен'}")
        b = InlineKeyboardBuilder()
        b.button(text="+7 дней", callback_data=f"a:grant:{uid}:7")
        b.button(text="+30 дней", callback_data=f"a:grant:{uid}:30")
        b.button(text="✅ Разблокировать" if u["banned"] else "🚫 Заблокировать", callback_data=f"a:ban:{uid}")
        b.button(text="← К списку", callback_data="a:users")
        b.adjust(2, 1, 1)
        return await edit(cb, text, kb_of(b))
    if act == "grant":
        uid, days = int(parts[2]), int(parts[3])
        until = grant_days(uid, days)
        await cb.answer(f"Подписка до {dt(until)}")
        try:
            await bot.send_message(uid, f"🎁 Вам начислено {days} дн. подписки. Активна до {dt(until)}.")
        except Exception:
            pass
        return await admin_screen(cb, ["a", "u", str(uid)])
    if act == "ban":
        uid = int(parts[2])
        DB.execute("UPDATE users SET banned=1-banned WHERE id=?", (uid,))
        DB.commit()
        return await admin_screen(cb, ["a", "u", str(uid)])
    if act in ("ok", "no"):
        p = settle(int(parts[2]), act == "ok")
        if not p:
            await cb.answer("Платёж уже обработан.", show_alert=True)
            return
        await cb.answer("Готово")
        if act == "ok":
            until = get_user(p["user_id"])["sub_until"]
            msg = f"✅ Оплата подтверждена. Подписка активна до {dt(until)}. Приятной работы!"
            kb = back_kb("search", "🔎 Найти клиентов")
        else:
            msg, kb = "❌ Платёж не подтверждён. Если вы оплатили, напишите администратору.", back_kb()
        try:
            await bot.send_message(p["user_id"], msg, reply_markup=kb)
        except Exception:
            pass
        verdict = "✅ подтверждён" if act == "ok" else "❌ отклонён"
        return await edit(cb, f"Платёж #{p['id']} {verdict}.", back)
    if act == "cast":
        WAIT[cb.from_user.id] = "a_cast"
        return await edit(cb, HINTS["a_cast"], back)
    if act == "castgo":
        text = CAST.pop(cb.from_user.id, "")
        if not text:
            return await edit(cb, "Текст рассылки пуст.", back)
        await edit(cb, "📣 Отправляю…", None)
        ok = fail = 0
        for r in DB.execute("SELECT id FROM users WHERE banned=0").fetchall():
            try:
                await bot.send_message(r["id"], text)
                ok += 1
            except Exception:
                fail += 1
            await asyncio.sleep(0.05)
        return await edit(cb, f"📣 Рассылка завершена. Доставлено: {ok}, не доставлено: {fail}.", back)
    if act == "set":
        b = InlineKeyboardBuilder()
        b.button(text=f"💰 Цена: {price()} ₽", callback_data="a:price")
        b.button(text=f"📆 Срок: {period()} дн.", callback_data="a:period")
        b.button(text=f"🎁 Пробный поиск: {'вкл' if trial_on() else 'выкл'}", callback_data="a:trial")
        b.button(text="← Админка", callback_data="a:home")
        b.adjust(1)
        return await edit(cb, "⚙️ <b>Настройки подписки</b>\n\nИзменения действуют для новых платежей.", kb_of(b))
    if act in ("price", "period"):
        WAIT[cb.from_user.id] = "a_" + act
        return await edit(cb, HINTS["a_" + act], back_kb("a:set", "← Назад"))
    if act == "trial":
        set_setting("trial", 0 if trial_on() else 1)
        return await admin_screen(cb, ["a", "set"])
    if act == "export":
        await cb.answer()
        wb = Workbook()
        ws = wb.active
        ws.title = "Платежи"
        ws.append(["№", "Дата", "Пользователь", "ID", "Сумма, ₽", "Дней", "Статус", "Подтверждён"])
        for r in DB.execute("SELECT p.*, u.username, u.name FROM payments p LEFT JOIN users u ON u.id=p.user_id "
                            "ORDER BY p.id DESC").fetchall():
            ws.append([r["id"], dt(r["created"], True), f"@{r['username']}" if r["username"] else r["name"],
                       r["user_id"], r["amount"], r["days"], STATUS[r["status"]], dt(r["confirmed"], True)])
        body(ws, 2, 22)
        header(ws, [7, 17, 22, 13, 11, 8, 16, 17])
        ws2 = wb.create_sheet("Пользователи")
        ws2.append(["ID", "Пользователь", "Регистрация", "Подписка до", "Поисков", "Пробный", "Заблокирован"])
        for r in DB.execute("SELECT * FROM users ORDER BY joined DESC").fetchall():
            ws2.append([r["id"], f"@{r['username']}" if r["username"] else r["name"], dt(r["joined"]),
                        dt(r["sub_until"]) if r["sub_until"] > now() else "нет", r["searches"],
                        "да" if r["trial_used"] else "нет", "да" if r["banned"] else "нет"])
        body(ws2, 2, 22)
        header(ws2, [13, 22, 14, 14, 9, 10, 14])
        path = os.path.join(tempfile.gettempdir(), "admin_export.xlsx")
        wb.save(path)
        return await cb.message.answer_document(FSInputFile(path, filename="payments_users.xlsx"))


@dp.callback_query(F.data.startswith("a:"))
async def cb_admin(cb: CallbackQuery):
    if cb.from_user.id not in ADMINS:
        return await cb.answer("Нет доступа.", show_alert=True)
    if cb.data.split(":")[1] not in ("ok", "no", "grant", "export"):
        await cb.answer()
    await admin_screen(cb, cb.data.split(":"))


# ───────────────────────── ввод текста ─────────────────────────
@dp.message(F.text)
async def on_text(m: Message):
    uid = m.from_user.id
    kind = WAIT.pop(uid, None)
    if not kind:
        return await m.answer(home_text(uid), reply_markup=home_kb(uid))
    if kind.startswith("a_"):
        if uid not in ADMINS:
            return
        back = InlineKeyboardMarkup(inline_keyboard=[[btn("← Админка", "a:home")]])
        if kind == "a_cast":
            CAST[uid] = m.text
            b = InlineKeyboardBuilder()
            b.button(text="📣 Отправить всем", callback_data="a:castgo")
            b.button(text="← Отмена", callback_data="a:home")
            b.adjust(1)
            return await m.answer("Предпросмотр:\n\n" + m.text, reply_markup=kb_of(b))
        try:
            v = int(m.text.strip())
            if v < (0 if kind == "a_price" else 1):
                raise ValueError
        except ValueError:
            WAIT[uid] = kind
            return await m.answer("Нужно целое число. " + HINTS[kind])
        set_setting("price" if kind == "a_price" else "period", v)
        return await m.answer("✅ Сохранено.", reply_markup=back)
    if not apply_text(uid, kind, m.text):
        WAIT[uid] = kind
        return await m.answer("Не подошло. " + HINTS[kind])
    await m.answer(summary(uid), reply_markup=main_kb(uid))


async def main():
    await bot.set_my_commands([BotCommand(command="menu", description="Главное меню"),
                               BotCommand(command="id", description="Узнать свой Telegram ID")])
    await dp.start_polling(bot)


if __name__ == "__main__":
    asyncio.run(main())
