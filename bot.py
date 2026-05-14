import os
import json
import logging
import asyncio
import re
from datetime import datetime
from pathlib import Path

import httpx
from telegram import Update, InlineKeyboardButton, InlineKeyboardMarkup
from telegram.ext import (
    Application, CommandHandler, MessageHandler,
    CallbackQueryHandler, ContextTypes, filters, ConversationHandler
)
import pdfplumber
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

# ── Config ───────────────────────────────────────────────────────────────────
BOT_TOKEN   = os.environ.get("BOT_TOKEN", "")
ADMIN_IDS   = [int(x) for x in os.environ.get("ADMIN_IDS", "").split(",") if x.strip()]
SALARY_PCT  = float(os.environ.get("SALARY_PCT", "20"))
BONUS_PCT   = float(os.environ.get("BONUS_PCT", "15"))
RATE_MARKUP = float(os.environ.get("RATE_MARKUP", "2"))   # % надбавка до курсу

DATA_DIR = Path("data")
DATA_DIR.mkdir(exist_ok=True)

# ── Load lookup ───────────────────────────────────────────────────────────────
with open("lookup.json", encoding="utf-8") as f:
    LOOKUP: dict = json.load(f)

# ── Persistent storage helpers ───────────────────────────────────────────────
def _user_file(user_id: int) -> Path:
    return DATA_DIR / f"{user_id}.json"

def load_user(user_id: int) -> dict:
    p = _user_file(user_id)
    if p.exists():
        return json.loads(p.read_text(encoding="utf-8"))
    return {"name": "", "invoices": []}

def save_user(user_id: int, data: dict):
    _user_file(user_id).write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")

# ── NBU rate fetch ────────────────────────────────────────────────────────────
async def get_nbu_rate(date_str: str) -> float | None:
    """date_str: DD.MM.YYYY"""
    try:
        dt = datetime.strptime(date_str, "%d.%m.%Y")
        url = f"https://bank.gov.ua/NBU_Exchange/exchange_site?start={dt:%Y%m%d}&end={dt:%Y%m%d}&valcode=EUR&sort=exchangedate&order=desc&json"
        async with httpx.AsyncClient(timeout=8) as client:
            r = await client.get(url)
            data = r.json()
            if data:
                rate = float(data[0]["rate"])
                return round(rate * (1 + RATE_MARKUP / 100), 4)
    except Exception as e:
        logger.warning(f"NBU rate error: {e}")
    return None

# ── PDF parser ────────────────────────────────────────────────────────────────
def parse_pdf(path: str) -> dict:
    """Extract invoice number, date, client, items from PDF."""
    result = {"invoice_num": "", "date": "", "client": "", "items": []}
    try:
        with pdfplumber.open(path) as pdf:
            text = "\n".join(p.extract_text() or "" for p in pdf.pages)
    except Exception as e:
        logger.error(f"PDF parse error: {e}")
        return result

    # Invoice number & date
    m = re.search(r"Рахунок на оплату\s*[№#]\s*(\d+)\s*від\s*([\d]+\s+\w+\s+\d{4})", text)
    if m:
        result["invoice_num"] = f"Рахунок №{m.group(1)} від {m.group(2)}"
        # parse date
        months = {"січня":1,"лютого":2,"березня":3,"квітня":4,"травня":5,"червня":6,
                  "липня":7,"серпня":8,"вересня":9,"жовтня":10,"листопада":11,"грудня":12}
        dm = re.search(r"(\d+)\s+(\w+)\s+(\d{4})", m.group(2))
        if dm:
            month_num = months.get(dm.group(2).lower(), 0)
            if month_num:
                result["date"] = f"{int(dm.group(1)):02d}.{month_num:02d}.{dm.group(3)}"

    # Client
    m2 = re.search(r"Покупець:\s*(.+?)(?:\n|Тел)", text, re.DOTALL)
    if m2:
        result["client"] = m2.group(1).strip()[:80]

    # Items — pattern: digit(s) + article-like code + qty + unit + price + total
    # Try table rows: number, name+article, qty, unit, price, sum
    item_pattern = re.compile(
        r"(\d{1,3})\s+"                          # row number
        r"[\w\s,\-\'\.]+?"                        # name (non-greedy)
        r"([A-Z0-9][A-Z0-9\-\/\.]{4,})\s+"       # article (starts uppercase, 5+ chars)
        r"(\d+)\s+шт\s+"                          # qty
        r"([\d\s]+[,.][\d]{2})\s+"               # price
        r"([\d\s]+[,.][\d]{2})"                   # total
    )
    seen = set()
    for m in item_pattern.finditer(text):
        art = m.group(2).strip()
        if art in seen:
            continue
        seen.add(art)
        qty = int(m.group(3))
        price_str = m.group(4).replace(" ", "").replace(",", ".")
        try:
            price = float(price_str)
        except:
            continue
        result["items"].append({"article": art, "qty": qty, "price_uah": price})

    # Fallback: simpler pattern if table extraction failed
    if not result["items"]:
        simple = re.compile(
            r"([A-Z0-9][A-Z0-9\-\/\.]{5,})\s+(\d{1,3})\s+шт\s+([\d\s]+[,.][\d]{2})"
        )
        for m in simple.finditer(text):
            art = m.group(1).strip()
            qty = int(m.group(2))
            price_str = m.group(3).replace(" ", "").replace(",", ".")
            try:
                price = float(price_str)
            except:
                continue
            result["items"].append({"article": art, "qty": qty, "price_uah": price})

    return result

# ── Lookup helper ─────────────────────────────────────────────────────────────
def lookup_article(art: str) -> dict:
    return LOOKUP.get(art, {})

# ── Excel generator ───────────────────────────────────────────────────────────
def build_excel(manager_name: str, invoices: list, month: str) -> str:
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = month

    thin = Side(style='thin', color='BDC3C7')
    brd = Border(left=thin, right=thin, top=thin, bottom=thin)

    # Title
    ws.merge_cells('A1:N1')
    ws['A1'] = f'РОЗРАХУНОК ЗП — {manager_name} — {month}'
    ws['A1'].font = Font(name='Arial', bold=True, size=13, color='FFFFFF')
    ws['A1'].fill = PatternFill('solid', fgColor='1F2D3D')
    ws['A1'].alignment = Alignment(horizontal='center', vertical='center')
    ws.row_dimensions[1].height = 26

    headers = ['Клієнт','Рахунок','Дата оплати','Артикул','Кть',
               'Закуп EUR','Мито%','Курс','Собів UAH','Ціна прод UAH',
               'Виторг UAH','Собів.загал','Прибуток','Склад?']
    ws.row_dimensions[2].height = 36
    for c, h in enumerate(headers, 1):
        cell = ws.cell(2, c, h)
        cell.font = Font(name='Arial', bold=True, size=8, color='FFFFFF')
        cell.fill = PatternFill('solid', fgColor='2C3E50')
        cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
        cell.border = brd

    col_widths = [18,22,12,22,5,10,7,10,12,14,13,13,13,8]
    for i, w in enumerate(col_widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w

    row = 3
    for inv in invoices:
        for item in inv.get("items", []):
            lu = lookup_article(item["article"])
            cost_eur = lu.get("cost_eur", 0)
            duty = lu.get("duty", 0.04)
            rate = item.get("rate", 52.0)
            cost_uah_unit = cost_eur * (1 + duty) * rate
            revenue = item["price_uah"] * item["qty"]
            cost_total = cost_uah_unit * item["qty"]
            is_stock = item.get("is_stock", False)

            if is_stock:
                # price from lookup as UAH
                price_list_uah = lu.get("cost_eur", 0) * rate
                profit = (item["price_uah"] - price_list_uah) * item["qty"]
            else:
                profit = revenue - cost_total

            vals = [
                inv.get("client",""),
                inv.get("invoice_num",""),
                inv.get("date",""),
                item["article"],
                item["qty"],
                cost_eur,
                f'{duty*100:.0f}%',
                rate,
                round(cost_uah_unit, 2),
                item["price_uah"],
                round(revenue, 2),
                round(cost_total, 2),
                round(profit, 2),
                "так" if is_stock else "",
            ]
            for c, v in enumerate(vals, 1):
                cell = ws.cell(row, c, v)
                cell.font = Font(name='Arial', size=9)
                cell.border = brd
                cell.alignment = Alignment(vertical='center')
                if c in [6,8,9,10,11,12,13]:
                    cell.number_format = '#,##0.00'
                if is_stock:
                    cell.fill = PatternFill('solid', fgColor='FDECEA')
                else:
                    cell.fill = PatternFill('solid', fgColor='F8F9FA' if row % 2 == 0 else 'FFFFFF')
            ws.row_dimensions[row].height = 17
            row += 1

    # Totals
    TOT = row
    ws.merge_cells(f'A{TOT}:J{TOT}')
    ws.cell(TOT, 1, 'ПІДСУМОК').font = Font(name='Arial', bold=True, size=10, color='FFFFFF')
    ws.cell(TOT, 1).fill = PatternFill('solid', fgColor='1A5276')
    ws.cell(TOT, 1).alignment = Alignment(horizontal='right')

    rev_col = get_column_letter(11)
    cost_col = get_column_letter(12)
    prof_col = get_column_letter(13)
    for c in [11, 12, 13]:
        col_l = get_column_letter(c)
        cell = ws.cell(TOT, c, f'=SUM({col_l}3:{col_l}{TOT-1})')
        cell.font = Font(name='Arial', bold=True, size=10, color='FFFFFF')
        cell.fill = PatternFill('solid', fgColor='1A5276')
        cell.number_format = '#,##0.00'
        cell.border = brd

    # Summary
    SR = TOT + 2
    summary = [
        ('Загальний виторг (грн):', f'={rev_col}{TOT}'),
        ('Загальна собівартість (грн):', f'={cost_col}{TOT}'),
        ('Валовий прибуток (грн):', f'={prof_col}{TOT}'),
        ('▶ Витрати на доставку (вручну, грн):', 0),
        ('Чистий прибуток (грн):', f'=D{SR+2}-D{SR+3}'),
        (f'ЗП ({SALARY_PCT}% від чистого прибутку):', f'=MAX(0,D{SR+4})*{SALARY_PCT}/100'),
    ]
    for i, (lbl, val) in enumerate(summary):
        r2 = SR + i
        ws.merge_cells(f'A{r2}:C{r2}')
        c_l = ws.cell(r2, 1, lbl)
        c_l.font = Font(name='Arial', size=10, bold=(i in [3,4,5]))
        c_l.alignment = Alignment(horizontal='right')
        c_v = ws.cell(r2, 4, val)
        c_v.number_format = '#,##0.00'
        if i == 3:
            c_v.fill = PatternFill('solid', fgColor='EBF5FB')
            c_v.font = Font(name='Arial', size=11, color='000080', bold=True)
        elif i == 5:
            c_v.fill = PatternFill('solid', fgColor='D5F5E3')
            c_v.font = Font(name='Arial', size=14, color='1E8449', bold=True)

    path = str(DATA_DIR / f"salary_{manager_name}_{month}.xlsx")
    wb.save(path)
    return path

# ── States ────────────────────────────────────────────────────────────────────
WAIT_NAME, WAIT_DATE, WAIT_STOCK, WAIT_DELIVERY = range(4)

# ── Handlers ──────────────────────────────────────────────────────────────────
async def start(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    user = load_user(update.effective_user.id)
    name = user.get("name", "")
    if name:
        await update.message.reply_text(
            f"👋 Привіт, {name}!\n\n"
            "Надішли PDF рахунку — я його оброблю автоматично.\n\n"
            "Команди:\n"
            "/звіт — отримати Excel з ЗП за місяць\n"
            "/очистити — почати місяць заново\n"
            "/імя — змінити ім'я\n"
            + ("/адмін — всі менеджери" if update.effective_user.id in ADMIN_IDS else "")
        )
    else:
        await update.message.reply_text(
            "👋 Привіт! Я бот для розрахунку ЗП.\n\n"
            "Як тебе звати? (введи своє ім'я)"
        )
        return WAIT_NAME

async def set_name(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    name = update.message.text.strip()
    user = load_user(update.effective_user.id)
    user["name"] = name
    save_user(update.effective_user.id, user)
    await update.message.reply_text(
        f"✅ Збережено! Привіт, {name}!\n\n"
        "Тепер надсилай PDF рахунки — я їх оброблятиму автоматично."
    )
    return ConversationHandler.END

async def handle_pdf(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    user_id = update.effective_user.id
    user = load_user(user_id)

    if not user.get("name"):
        await update.message.reply_text("Спочатку введи своє ім'я: /start")
        return

    msg = await update.message.reply_text("⏳ Обробляю PDF...")

    # Download PDF
    doc = update.message.document
    file = await ctx.bot.get_file(doc.file_id)
    pdf_path = str(DATA_DIR / f"{user_id}_{doc.file_id}.pdf")
    await file.download_to_drive(pdf_path)

    # Parse
    parsed = parse_pdf(pdf_path)
    os.remove(pdf_path)

    if not parsed["items"]:
        await msg.edit_text("❌ Не вдалось розпізнати позиції в PDF. Перевір формат файлу.")
        return

    # Lookup articles
    found, not_found = [], []
    for item in parsed["items"]:
        lu = lookup_article(item["article"])
        if lu.get("cost_eur"):
            item.update(lu)
            found.append(item["article"])
        else:
            not_found.append(item["article"])

    # Store pending invoice in context
    ctx.user_data["pending_invoice"] = {
        "invoice_num": parsed["invoice_num"],
        "date": parsed["date"],
        "client": parsed["client"],
        "items": parsed["items"],
    }

    # Build preview
    lines = [f"📄 *{parsed['invoice_num']}*", f"👤 {parsed['client']}", ""]
    for item in parsed["items"]:
        lu = lookup_article(item["article"])
        cost = f"{lu['cost_eur']:.2f} EUR" if lu.get("cost_eur") else "❓ немає в довіднику"
        lines.append(f"• `{item['article']}` × {item['qty']} шт — {item['price_uah']:,.2f} грн/шт | Закуп: {cost}")

    if not_found:
        lines.append(f"\n⚠️ Не знайдено в довіднику ({len(not_found)}):")
        for a in not_found:
            lines.append(f"  `{a}`")

    lines.append(f"\n✅ Розпізнано: {len(found)}/{len(parsed['items'])} позицій")
    lines.append("\nВведи дату оплати клієнтом (формат: ДД.ММ.РРРР):")

    await msg.edit_text("\n".join(lines), parse_mode="Markdown")
    return WAIT_DATE

async def handle_date(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    date_str = update.message.text.strip()
    if not re.match(r"\d{2}\.\d{2}\.\d{4}", date_str):
        await update.message.reply_text("❌ Формат: ДД.ММ.РРРР (наприклад 23.04.2026)")
        return WAIT_DATE

    ctx.user_data["pending_invoice"]["date"] = date_str

    # Fetch NBU rate
    rate = await get_nbu_rate(date_str)
    if rate:
        ctx.user_data["pending_invoice"]["rate"] = rate
        ctx.user_data["pending_invoice"]["items"] = [
            {**item, "rate": rate} for item in ctx.user_data["pending_invoice"]["items"]
        ]
        rate_msg = f"💱 Курс НБУ на {date_str}: *{rate:.4f} грн/EUR* (міжбанк +{RATE_MARKUP}%)"
    else:
        ctx.user_data["pending_invoice"]["rate"] = 52.0
        rate_msg = f"⚠️ Не вдалось отримати курс на {date_str}. Використовую 52.00"

    await update.message.reply_text(
        rate_msg + "\n\nЄ складські товари (червоні)? Введи артикули через кому або надішли /ні",
        parse_mode="Markdown"
    )
    return WAIT_STOCK

async def handle_stock(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    text = update.message.text.strip()
    inv = ctx.user_data.get("pending_invoice", {})

    if text.lower() not in ["/ні", "ні", "нет", "no", "/no"]:
        stock_arts = [a.strip() for a in text.replace("/", "").split(",")]
        for item in inv.get("items", []):
            if item["article"] in stock_arts:
                item["is_stock"] = True

    # Save invoice
    user_id = update.effective_user.id
    user = load_user(user_id)
    user.setdefault("invoices", []).append(inv)
    save_user(user_id, user)
    ctx.user_data.pop("pending_invoice", None)

    # Quick profit preview
    total_profit = 0
    for item in inv.get("items", []):
        lu = lookup_article(item["article"])
        cost_eur = lu.get("cost_eur", 0)
        duty = lu.get("duty", 0.04)
        rate = item.get("rate", 52.0)
        cost_uah = cost_eur * (1 + duty) * rate * item["qty"]
        revenue = item["price_uah"] * item["qty"]
        if item.get("is_stock"):
            price_list_uah = cost_eur * rate
            profit = (item["price_uah"] - price_list_uah) * item["qty"]
        else:
            profit = revenue - cost_uah
        total_profit += profit

    total_invoices = len(user.get("invoices", []))
    await update.message.reply_text(
        f"✅ *Рахунок збережено!*\n\n"
        f"📊 Прибуток по цьому рахунку: *{total_profit:,.0f} грн*\n"
        f"📁 Всього рахунків цього місяця: {total_invoices}\n\n"
        f"Надішли наступний PDF або /звіт для отримання Excel.",
        parse_mode="Markdown"
    )
    return ConversationHandler.END

async def cmd_report(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    user_id = update.effective_user.id
    user = load_user(user_id)
    invoices = user.get("invoices", [])

    if not invoices:
        await update.message.reply_text("📭 Немає рахунків. Надішли PDF рахунок спочатку.")
        return

    await update.message.reply_text("⏳ Генерую Excel...")

    # Ask for delivery costs
    ctx.user_data["report_mode"] = True
    await update.message.reply_text(
        "Введи витрати на доставку за місяць (грн), або /0 якщо немає:"
    )
    return WAIT_DELIVERY

async def handle_delivery(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    text = update.message.text.strip().replace("/", "")
    try:
        delivery = float(text.replace(",", ".").replace(" ", ""))
    except:
        delivery = 0

    user_id = update.effective_user.id
    user = load_user(user_id)
    invoices = user.get("invoices", [])
    name = user.get("name", str(user_id))
    month = datetime.now().strftime("%m.%Y")

    # Calculate totals
    total_revenue = sum(
        item["price_uah"] * item["qty"]
        for inv in invoices for item in inv.get("items", [])
    )
    total_profit = 0
    for inv in invoices:
        for item in inv.get("items", []):
            lu = lookup_article(item["article"])
            cost_eur = lu.get("cost_eur", 0)
            duty = lu.get("duty", 0.04)
            rate = item.get("rate", inv.get("rate", 52.0))
            cost_uah = cost_eur * (1 + duty) * rate * item["qty"]
            revenue = item["price_uah"] * item["qty"]
            if item.get("is_stock"):
                price_list_uah = cost_eur * rate
                profit = (item["price_uah"] - price_list_uah) * item["qty"]
            else:
                profit = revenue - cost_uah
            total_profit += profit

    net_profit = total_profit - delivery
    salary = max(0, net_profit) * SALARY_PCT / 100

    # Build Excel
    path = build_excel(name, invoices, month)

    await update.message.reply_document(
        document=open(path, "rb"),
        filename=f"ЗП_{name}_{month}.xlsx",
        caption=(
            f"📊 *Звіт за {month}*\n\n"
            f"Виторг: *{total_revenue:,.0f} грн*\n"
            f"Прибуток: *{total_profit:,.0f} грн*\n"
            f"Доставка: *{delivery:,.0f} грн*\n"
            f"Чистий прибуток: *{net_profit:,.0f} грн*\n"
            f"━━━━━━━━━━━━━\n"
            f"💰 *ЗП: {salary:,.0f} грн*"
        ),
        parse_mode="Markdown"
    )
    return ConversationHandler.END

async def cmd_clear(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    keyboard = [[
        InlineKeyboardButton("✅ Так, очистити", callback_data="clear_yes"),
        InlineKeyboardButton("❌ Скасувати", callback_data="clear_no"),
    ]]
    await update.message.reply_text(
        "⚠️ Видалити всі рахунки поточного місяця?",
        reply_markup=InlineKeyboardMarkup(keyboard)
    )

async def callback_clear(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    if query.data == "clear_yes":
        user = load_user(query.from_user.id)
        user["invoices"] = []
        save_user(query.from_user.id, user)
        await query.edit_message_text("✅ Рахунки очищено. Починай новий місяць!")
    else:
        await query.edit_message_text("Скасовано.")

async def cmd_admin(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    if update.effective_user.id not in ADMIN_IDS:
        await update.message.reply_text("⛔ Немає доступу.")
        return

    all_users = list(DATA_DIR.glob("*.json"))
    if not all_users:
        await update.message.reply_text("Немає даних.")
        return

    lines = ["📋 *Всі менеджери:*\n"]
    for f in all_users:
        try:
            data = json.loads(f.read_text(encoding="utf-8"))
            name = data.get("name", f.stem)
            count = len(data.get("invoices", []))
            total_items = sum(len(inv.get("items",[])) for inv in data.get("invoices",[]))

            total_profit = 0
            for inv in data.get("invoices", []):
                for item in inv.get("items", []):
                    lu = lookup_article(item["article"])
                    cost_eur = lu.get("cost_eur", 0)
                    duty = lu.get("duty", 0.04)
                    rate = item.get("rate", 52.0)
                    cost_uah = cost_eur * (1 + duty) * rate * item["qty"]
                    revenue = item["price_uah"] * item["qty"]
                    if item.get("is_stock"):
                        price_list_uah = cost_eur * rate
                        profit = (item["price_uah"] - price_list_uah) * item["qty"]
                    else:
                        profit = revenue - cost_uah
                    total_profit += profit

            salary = max(0, total_profit) * SALARY_PCT / 100
            lines.append(f"👤 *{name}*: {count} рахунків, {total_items} позицій\n   Прибуток: {total_profit:,.0f} грн | ЗП: {salary:,.0f} грн\n")
        except:
            pass

    await update.message.reply_text("\n".join(lines), parse_mode="Markdown")

async def cmd_name(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text("Введи нове ім'я:")
    return WAIT_NAME

# ── Main ──────────────────────────────────────────────────────────────────────
def main():
    app = Application.builder().token(BOT_TOKEN).build()

    conv = ConversationHandler(
        entry_points=[
            CommandHandler("start", start),
            MessageHandler(filters.Document.PDF, handle_pdf),
            CommandHandler("звіт", cmd_report),
            CommandHandler("report", cmd_report),
            CommandHandler("імя", cmd_name),
        ],
        states={
            WAIT_NAME:     [MessageHandler(filters.TEXT & ~filters.COMMAND, set_name)],
            WAIT_DATE:     [MessageHandler(filters.TEXT & ~filters.COMMAND, handle_date)],
            WAIT_STOCK:    [MessageHandler(filters.TEXT, handle_stock),
                            CommandHandler("ні", handle_stock)],
            WAIT_DELIVERY: [MessageHandler(filters.TEXT, handle_delivery),
                            CommandHandler("0", handle_delivery)],
        },
        fallbacks=[CommandHandler("start", start)],
        allow_reentry=True,
    )

    app.add_handler(conv)
    app.add_handler(CommandHandler("очистити", cmd_clear))
    app.add_handler(CommandHandler("clear", cmd_clear))
    app.add_handler(CommandHandler("адмін", cmd_admin))
    app.add_handler(CommandHandler("admin", cmd_admin))
    app.add_handler(CallbackQueryHandler(callback_clear, pattern="^clear_"))

    logger.info("Bot started")
    app.run_polling(drop_pending_updates=True)

if __name__ == "__main__":
    main()
