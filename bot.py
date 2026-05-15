import os
import json
import logging
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
import gspread
from google.oauth2.service_account import Credentials

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

BOT_TOKEN    = os.environ.get("BOT_TOKEN", "")
ADMIN_IDS    = [int(x) for x in os.environ.get("ADMIN_IDS", "").split(",") if x.strip()]
SALARY_PCT   = float(os.environ.get("SALARY_PCT", "20"))
RATE_MARKUP  = float(os.environ.get("RATE_MARKUP", "2"))
SHEET_ID     = os.environ.get("SHEET_ID", "")
GOOGLE_CREDS = os.environ.get("GOOGLE_CREDS", "")

DATA_DIR = Path("data")
DATA_DIR.mkdir(exist_ok=True)

with open("lookup.json", encoding="utf-8") as f:
    LOOKUP: dict = json.load(f)

ORDERS_SHEET_ID = os.environ.get("ORDERS_SHEET_ID", "")
_live_cache: dict = {}
_cache_time = None

def refresh_live_lookup():
    global _live_cache, _cache_time
    import time
    now = time.time()
    # Refresh every 30 minutes
    if _cache_time and now - _cache_time < 1800:
        return
    try:
        creds_dict = json.loads(GOOGLE_CREDS)
        scopes = ["https://spreadsheets.google.com/feeds",
                  "https://www.googleapis.com/auth/drive"]
        creds = Credentials.from_service_account_info(creds_dict, scopes=scopes)
        gc = gspread.authorize(creds)
        sh = gc.open_by_key(ORDERS_SHEET_ID)
        result = {}
        for ws in sh.worksheets():
            src = ws.title
            rows = ws.get_all_values()
            if not rows: continue
            headers = [h.strip() for h in rows[0]]
            # Find column indices
            def col(name):
                for i,h in enumerate(headers):
                    if name.lower() in h.lower(): return i
                return -1
            art_col = col("артикул")
            uktved_col = col("код товару")
            duty_col = col("мито")
            price_col = col("ціна за одиницю")
            weight_col = col("нетто за 1")
            brand_col = col("виробник")
            if art_col < 0: continue
            for row in rows[1:]:
                if len(row) <= art_col: continue
                art = str(row[art_col]).strip()
                if not art: continue
                try: price = float(str(row[price_col]).replace(",",".").replace(" ","")) if price_col>=0 and price_col<len(row) else 0
                except: price = 0
                try: duty = float(str(row[duty_col]).replace(",",".").replace(" ","").replace("%","")) if duty_col>=0 and duty_col<len(row) else 0.04
                except: duty = 0.04
                if duty > 1: duty = duty / 100
                try: weight = float(str(row[weight_col]).replace(",",".").replace(" ","")) if weight_col>=0 and weight_col<len(row) else 0
                except: weight = 0
                result[art] = {
                    "uktved": str(row[uktved_col]).strip() if uktved_col>=0 and uktved_col<len(row) else "",
                    "duty": duty,
                    "cost_eur": price,
                    "weight": weight,
                    "brand": str(row[brand_col]).strip() if brand_col>=0 and brand_col<len(row) else "",
                    "source": src,
                }
        _live_cache = result
        _cache_time = now
        logger.info(f"Live lookup refreshed: {len(result)} articles")
    except Exception as e:
        logger.error(f"Live lookup error: {e}")

# ── Google Sheets ─────────────────────────────────────────────────────────────
def get_gsheet():
    try:
        creds_dict = json.loads(GOOGLE_CREDS)
        scopes = ["https://spreadsheets.google.com/feeds",
                  "https://www.googleapis.com/auth/drive"]
        creds = Credentials.from_service_account_info(creds_dict, scopes=scopes)
        gc = gspread.authorize(creds)
        return gc.open_by_key(SHEET_ID)
    except Exception as e:
        logger.error(f"GSheet connect error: {e}")
        return None

def get_or_create_ws(spreadsheet, name: str):
    try:
        return spreadsheet.worksheet(name)
    except gspread.WorksheetNotFound:
        ws = spreadsheet.add_worksheet(title=name, rows=500, cols=20)
        headers = ["Менеджер","Клієнт","Рахунок","Дата оплати","Артикул",
                   "Кть","Закуп EUR","Мито%","Курс","Собів UAH/шт","Собів загал",
                   "Ціна прод UAH","Виторг","Прибуток","Склад?",
                   "УКТЗЕД","Бренд","Джерело","Додано"]
        ws.append_row(headers)
        ws.format("A1:S1", {
            "backgroundColor": {"red":0.17,"green":0.18,"blue":0.24},
            "textFormat": {"bold":True,"foregroundColor":{"red":1,"green":1,"blue":1}},
            "horizontalAlignment": "CENTER"
        })
        return ws

def append_to_sheets(manager_name: str, inv: dict):
    try:
        sh = get_gsheet()
        if not sh:
            return False
        ws_mgr = get_or_create_ws(sh, manager_name)
        ws_all = get_or_create_ws(sh, "ВСІ")
        rows_added = []
        for item in inv.get("items", []):
            lu = lookup_article(item["article"])
            cost_eur = lu.get("cost_eur", 0)
            duty = lu.get("duty", 0.04)
            rate = item.get("rate", inv.get("rate", 52.0))
            cost_unit = cost_eur * (1 + duty) * rate
            cost_total = cost_unit * item["qty"]
            revenue = item["price_uah"] * item["qty"]
            if item.get("is_stock"):
                profit = (item["price_uah"] - cost_eur * rate) * item["qty"]
            else:
                profit = revenue - cost_total
            row = [
                manager_name, inv.get("client",""), inv.get("invoice_num",""),
                inv.get("date",""), item["article"], item["qty"],
                round(cost_eur,2), f'{duty*100:.0f}%', round(rate,4),
                round(cost_unit,2), round(cost_total,2),
                item["price_uah"], round(revenue,2), round(profit,2),
                "так" if item.get("is_stock") else "",
                lu.get("uktved",""), lu.get("brand",""), lu.get("source",""),
                datetime.now().strftime("%d.%m.%Y %H:%M"),
            ]
            rows_added.append(row)
            ws_mgr.append_row(row)
            ws_all.append_row(row)

        # Color red stock rows
        all_vals = ws_mgr.get_all_values()
        for i, r in enumerate(all_vals[1:], 2):
            if len(r) >= 15 and r[14] == "так":
                ws_mgr.format(f"A{i}:S{i}", {
                    "backgroundColor": {"red":0.99,"green":0.87,"blue":0.87}
                })
        return True
    except Exception as e:
        logger.error(f"Sheet append error: {e}")
        return False

# ── Storage ───────────────────────────────────────────────────────────────────
def _uf(uid): return DATA_DIR / f"{uid}.json"
def load_user(uid):
    p = _uf(uid)
    return json.loads(p.read_text(encoding="utf-8")) if p.exists() else {"name":"","invoices":[]}
def save_user(uid, data):
    _uf(uid).write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")

# ── NBU rate ──────────────────────────────────────────────────────────────────
async def get_nbu_rate(date_str: str):
    try:
        dt = datetime.strptime(date_str, "%d.%m.%Y")
        url = f"https://bank.gov.ua/NBU_Exchange/exchange_site?start={dt:%Y%m%d}&end={dt:%Y%m%d}&valcode=EUR&sort=exchangedate&order=desc&json"
        async with httpx.AsyncClient(timeout=8) as client:
            r = await client.get(url)
            data = r.json()
            if data:
                return round(float(data[0]["rate"]) * (1 + RATE_MARKUP / 100), 4)
    except Exception as e:
        logger.warning(f"NBU error: {e}")
    return None

# ── PDF parser ────────────────────────────────────────────────────────────────
def parse_pdf(path: str) -> dict:
    result = {"invoice_num":"","date":"","client":"","items":[]}
    try:
        with pdfplumber.open(path) as pdf:
            text = "\n".join(p.extract_text() or "" for p in pdf.pages)
    except Exception as e:
        logger.error(f"PDF error: {e}")
        return result

    m = re.search(r"Рахунок на оплату\s*[№#]\s*(\d+)\s*від\s*([\d]+\s+\w+\s+\d{4})", text)
    if m:
        result["invoice_num"] = f"Рахунок №{m.group(1)} від {m.group(2)}"
        months = {"січня":1,"лютого":2,"березня":3,"квітня":4,"травня":5,"червня":6,
                  "липня":7,"серпня":8,"вересня":9,"жовтня":10,"листопада":11,"грудня":12}
        dm = re.search(r"(\d+)\s+(\w+)\s+(\d{4})", m.group(2))
        if dm:
            mn = months.get(dm.group(2).lower(), 0)
            if mn:
                result["date"] = f"{int(dm.group(1)):02d}.{mn:02d}.{dm.group(3)}"

    m2 = re.search(r"Покупець:\s*(.+?)(?:\n|Тел)", text, re.DOTALL)
    if m2:
        result["client"] = m2.group(1).strip()[:80]

    item_pat = re.compile(
        r"\d{1,3}\s+[\w\s,\-\'\.]+?([A-Z0-9][A-Z0-9\-\/\.]{4,})\s+(\d+)\s+шт\s+([\d\s]+[,.][\d]{2})\s+([\d\s]+[,.][\d]{2})"
    )
    seen = set()
    for m in item_pat.finditer(text):
        art = m.group(1).strip()
        if art in seen: continue
        seen.add(art)
        try:
            price = float(m.group(3).replace(" ","").replace(",","."))
            result["items"].append({"article":art,"qty":int(m.group(2)),"price_uah":price})
        except: pass

    if not result["items"]:
        for m in re.finditer(r"([A-Z0-9][A-Z0-9\-\/\.]{5,})\s+(\d{1,3})\s+шт\s+([\d\s]+[,.][\d]{2})", text):
            try:
                result["items"].append({"article":m.group(1).strip(),"qty":int(m.group(2)),
                                        "price_uah":float(m.group(3).replace(" ","").replace(",","."))})
            except: pass
    return result

def lookup_article(art):
    # Try live cache first, fallback to static lookup.json
    if ORDERS_SHEET_ID and _live_cache:
        return _live_cache.get(art, LOOKUP.get(art, {}))
    return LOOKUP.get(art, {})

# ── Excel builder ─────────────────────────────────────────────────────────────
def build_excel(manager_name, invoices, month):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = month
    thin = Side(style='thin', color='BDC3C7')
    brd = Border(left=thin,right=thin,top=thin,bottom=thin)

    ws.merge_cells('A1:N1')
    ws['A1'] = f'РОЗРАХУНОК ЗП — {manager_name} — {month}'
    ws['A1'].font = Font(name='Arial',bold=True,size=13,color='FFFFFF')
    ws['A1'].fill = PatternFill('solid',fgColor='1F2D3D')
    ws['A1'].alignment = Alignment(horizontal='center',vertical='center')
    ws.row_dimensions[1].height = 26

    headers = ['Клієнт','Рахунок','Дата','Артикул','Кть','Закуп EUR','Мито%',
               'Курс','Собів UAH','Ціна UAH','Виторг','Собів загал','Прибуток','Склад?']
    ws.row_dimensions[2].height = 36
    for c,h in enumerate(headers,1):
        cell = ws.cell(2,c,h)
        cell.font = Font(name='Arial',bold=True,size=8,color='FFFFFF')
        cell.fill = PatternFill('solid',fgColor='2C3E50')
        cell.alignment = Alignment(horizontal='center',vertical='center',wrap_text=True)
        cell.border = brd

    for i,w in enumerate([18,22,12,22,5,10,7,10,12,14,13,13,13,8],1):
        ws.column_dimensions[get_column_letter(i)].width = w

    row = 3
    for inv in invoices:
        for item in inv.get("items",[]):
            lu = lookup_article(item["article"])
            cost_eur = lu.get("cost_eur",0)
            duty = lu.get("duty",0.04)
            rate = item.get("rate", inv.get("rate",52.0))
            cost_unit = cost_eur*(1+duty)*rate
            revenue = item["price_uah"]*item["qty"]
            cost_total = cost_unit*item["qty"]
            profit = (item["price_uah"]-cost_eur*rate)*item["qty"] if item.get("is_stock") else revenue-cost_total
            is_stock = item.get("is_stock",False)
            vals = [inv.get("client",""),inv.get("invoice_num",""),inv.get("date",""),
                    item["article"],item["qty"],cost_eur,f'{duty*100:.0f}%',rate,
                    round(cost_unit,2),item["price_uah"],round(revenue,2),round(cost_total,2),round(profit,2),
                    "так" if is_stock else ""]
            for c,v in enumerate(vals,1):
                cell = ws.cell(row,c,v)
                cell.font = Font(name='Arial',size=9)
                cell.border = brd
                cell.alignment = Alignment(vertical='center')
                if c in [6,8,9,10,11,12,13]: cell.number_format = '#,##0.00'
                cell.fill = PatternFill('solid',fgColor='FDECEA' if is_stock else ('F8F9FA' if row%2==0 else 'FFFFFF'))
            ws.row_dimensions[row].height = 17
            row += 1

    TOT = row
    ws.merge_cells(f'A{TOT}:J{TOT}')
    ws.cell(TOT,1,'ПІДСУМОК').font = Font(name='Arial',bold=True,size=10,color='FFFFFF')
    ws.cell(TOT,1).fill = PatternFill('solid',fgColor='1A5276')
    ws.cell(TOT,1).alignment = Alignment(horizontal='right')
    for c in [11,12,13]:
        cl = get_column_letter(c)
        cell = ws.cell(TOT,c,f'=SUM({cl}3:{cl}{TOT-1})')
        cell.font = Font(name='Arial',bold=True,size=10,color='FFFFFF')
        cell.fill = PatternFill('solid',fgColor='1A5276')
        cell.number_format = '#,##0.00'
        cell.border = brd

    SR = TOT+2
    rc = get_column_letter(13)
    labels = [('Валовий прибуток (грн):',f'={rc}{TOT}'),
              ('▶ Витрати на доставку (вручну):',0),
              ('Чистий прибуток (грн):',f'=D{SR}-D{SR+1}'),
              (f'ЗП ({SALARY_PCT}% від чистого прибутку):',f'=MAX(0,D{SR+2})*{SALARY_PCT}/100')]
    for i,(lbl,val) in enumerate(labels):
        r2=SR+i
        ws.merge_cells(f'A{r2}:C{r2}')
        ws.cell(r2,1,lbl).font=Font(name='Arial',size=10,bold=(i in[1,2,3]))
        ws.cell(r2,1).alignment=Alignment(horizontal='right')
        cv=ws.cell(r2,4,val)
        cv.number_format='#,##0.00'
        if i==1:
            cv.fill=PatternFill('solid',fgColor='EBF5FB')
            cv.font=Font(name='Arial',size=11,color='000080',bold=True)
        elif i==3:
            cv.fill=PatternFill('solid',fgColor='D5F5E3')
            cv.font=Font(name='Arial',size=14,color='1E8449',bold=True)

    path = str(DATA_DIR/f"salary_{manager_name}_{month}.xlsx")
    wb.save(path)
    return path

# ── States ────────────────────────────────────────────────────────────────────
WAIT_NAME, WAIT_DATE, WAIT_STOCK, WAIT_DELIVERY = range(4)

# ── Handlers ──────────────────────────────────────────────────────────────────
async def start(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    user = load_user(update.effective_user.id)
    name = user.get("name","")
    if name:
        await update.message.reply_text(
            f"👋 Привіт, {name}!\n\n"
            "Надішли PDF рахунку — я його оброблю.\n\n"
            "Команди:\n"
            "/report — Excel з ЗП за місяць\n"
            "/clear — очистити місяць\n"
            "/name — змінити ім'я"
            + ("\n/admin — всі менеджери" if update.effective_user.id in ADMIN_IDS else "")
        )
    else:
        await update.message.reply_text("👋 Привіт! Як тебе звати? (введи ім'я)")
        return WAIT_NAME

async def set_name(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    name = update.message.text.strip()
    user = load_user(update.effective_user.id)
    user["name"] = name
    save_user(update.effective_user.id, user)
    await update.message.reply_text(f"✅ Збережено! Привіт, {name}!\n\nНадсилай PDF рахунки.")
    return ConversationHandler.END

async def handle_pdf(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    uid = update.effective_user.id
    user = load_user(uid)
    if not user.get("name"):
        await update.message.reply_text("Спочатку введи ім'я: /start")
        return

    refresh_live_lookup()
    msg = await update.message.reply_text("⏳ Обробляю PDF...")
    doc = update.message.document
    file = await ctx.bot.get_file(doc.file_id)
    pdf_path = str(DATA_DIR/f"{uid}_{doc.file_id}.pdf")
    await file.download_to_drive(pdf_path)

    parsed = parse_pdf(pdf_path)
    os.remove(pdf_path)

    if not parsed["items"]:
        await msg.edit_text("❌ Не вдалось розпізнати позиції в PDF.")
        return

    found, not_found = [], []
    for item in parsed["items"]:
        lu = lookup_article(item["article"])
        if lu.get("cost_eur"):
            item.update(lu)
            found.append(item["article"])
        else:
            not_found.append(item["article"])

    ctx.user_data["pending_invoice"] = {
        "invoice_num": parsed["invoice_num"],
        "date": parsed["date"],
        "client": parsed["client"],
        "items": parsed["items"],
    }

    lines = [f"📄 *{parsed['invoice_num']}*", f"👤 {parsed['client']}", ""]
    for item in parsed["items"]:
        lu = lookup_article(item["article"])
        cost = f"{lu['cost_eur']:.2f} EUR" if lu.get("cost_eur") else "❓ немає в довіднику"
        lines.append(f"• `{item['article']}` × {item['qty']} — {item['price_uah']:,.2f} грн | Закуп: {cost}")

    if not_found:
        lines.append(f"\n⚠️ Не знайдено ({len(not_found)}): {', '.join(not_found[:5])}")

    lines.append(f"\n✅ Знайдено: {len(found)}/{len(parsed['items'])}")
    lines.append("\nВведи дату оплати (ДД.ММ.РРРР):")

    await msg.edit_text("\n".join(lines), parse_mode="Markdown")
    return WAIT_DATE

async def handle_date(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    date_str = update.message.text.strip()
    if not re.match(r"\d{2}\.\d{2}\.\d{4}", date_str):
        await update.message.reply_text("❌ Формат: ДД.ММ.РРРР (наприклад 23.04.2026)")
        return WAIT_DATE

    ctx.user_data["pending_invoice"]["date"] = date_str
    rate = await get_nbu_rate(date_str)
    if rate:
        ctx.user_data["pending_invoice"]["rate"] = rate
        for item in ctx.user_data["pending_invoice"]["items"]:
            item["rate"] = rate
        rate_msg = f"💱 Курс НБУ на {date_str}: *{rate:.4f} грн/EUR* (+{RATE_MARKUP}%)"
    else:
        ctx.user_data["pending_invoice"]["rate"] = 52.0
        rate_msg = f"⚠️ Не вдалось отримати курс. Використовую 52.00"

    inv = ctx.user_data.get("pending_invoice", {})
    items = inv.get("items", [])
    ctx.user_data["stock_selected"] = set()
    await update.message.reply_text(rate_msg, parse_mode="Markdown")
    await update.message.reply_text(
        "📦 *Вибери складські товари* (натисни щоб відмітити):",
        reply_markup=build_stock_keyboard(items, set()),
        parse_mode="Markdown"
    )
    return WAIT_STOCK

def build_stock_keyboard(items: list, selected: set) -> InlineKeyboardMarkup:
    """Build keyboard with checkboxes for each item."""
    buttons = []
    for item in items:
        art = item["article"]
        qty = item["qty"]
        check = "✅" if art in selected else "☐"
        buttons.append([InlineKeyboardButton(
            f"{check} {art} × {qty}",
            callback_data=f"stock_{art}"
        )])
    buttons.append([
        InlineKeyboardButton("✅ Готово — немає складських", callback_data="stock_done_none"),
        InlineKeyboardButton("💾 Зберегти", callback_data="stock_done"),
    ])
    return InlineKeyboardMarkup(buttons)

async def handle_stock(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    inv = ctx.user_data.get("pending_invoice", {})
    items = inv.get("items", [])
    ctx.user_data["stock_selected"] = set()

    await update.message.reply_text(
        "📦 *Вибери складські товари* (натисни щоб відмітити):",
        reply_markup=build_stock_keyboard(items, set()),
        parse_mode="Markdown"
    )
    return WAIT_STOCK

async def callback_stock(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()

    if query.data == "stock_done_none":
        # No stock items
        ctx.user_data["stock_selected"] = set()
        await save_invoice(query, ctx)
        return ConversationHandler.END

    if query.data == "stock_done":
        await save_invoice(query, ctx)
        return ConversationHandler.END

    if query.data.startswith("stock_"):
        art = query.data[6:]
        selected = ctx.user_data.get("stock_selected", set())
        if art in selected:
            selected.discard(art)
        else:
            selected.add(art)
        ctx.user_data["stock_selected"] = selected

        inv = ctx.user_data.get("pending_invoice", {})
        await query.edit_message_reply_markup(
            reply_markup=build_stock_keyboard(inv.get("items", []), selected)
        )

async def save_invoice(query, ctx: ContextTypes.DEFAULT_TYPE):
    uid = query.from_user.id
    user = load_user(uid)
    inv = ctx.user_data.get("pending_invoice", {})
    selected = ctx.user_data.get("stock_selected", set())

    # Mark stock items
    for item in inv.get("items", []):
        item["is_stock"] = item["article"] in selected

    user.setdefault("invoices", []).append(inv)
    save_user(uid, user)

    manager_name = user.get("name", str(uid))
    sheet_ok = append_to_sheets(manager_name, inv)
    sheet_msg = "📊 Записано в Google Sheets ✓" if sheet_ok else "⚠️ Google Sheets недоступний"

    ctx.user_data.pop("pending_invoice", None)
    ctx.user_data.pop("stock_selected", None)

    total_profit = 0
    for item in inv.get("items", []):
        lu = lookup_article(item["article"])
        cost_eur = lu.get("cost_eur", 0)
        duty = lu.get("duty", 0.04)
        rate = item.get("rate", 52.0)
        cost_uah = cost_eur * (1 + duty) * rate * item["qty"]
        revenue = item["price_uah"] * item["qty"]
        profit = (item["price_uah"] - cost_eur * rate) * item["qty"] if item.get("is_stock") else revenue - cost_uah
        total_profit += profit

    stock_count = len(selected)
    stock_msg = f"🔴 Складських: {stock_count}" if stock_count else "🟢 Складських немає"

    await query.edit_message_text(
        f"✅ *Рахунок збережено!*\n\n"
        f"💰 Прибуток: *{total_profit:,.0f} грн*\n"
        f"{stock_msg}\n"
        f"📁 Рахунків цього місяця: {len(user.get('invoices', []))}\n"
        f"{sheet_msg}\n\n"
        f"Надішли наступний PDF або /report для Excel.",
        parse_mode="Markdown"
    )

async def cmd_report(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    uid = update.effective_user.id
    user = load_user(uid)
    if not user.get("invoices"):
        await update.message.reply_text("📭 Немає рахунків цього місяця.")
        return
    await update.message.reply_text("Введи витрати на доставку (грн) або /0 якщо немає:")
    return WAIT_DELIVERY

async def handle_delivery(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    text = update.message.text.strip().replace("/","")
    try:
        delivery = float(text.replace(",",".").replace(" ",""))
    except:
        delivery = 0

    uid = update.effective_user.id
    user = load_user(uid)
    invoices = user.get("invoices",[])
    name = user.get("name", str(uid))
    month = datetime.now().strftime("%m.%Y")

    total_profit = 0
    total_revenue = 0
    for inv in invoices:
        for item in inv.get("items",[]):
            lu = lookup_article(item["article"])
            cost_eur = lu.get("cost_eur",0)
            duty = lu.get("duty",0.04)
            rate = item.get("rate", inv.get("rate",52.0))
            cost_uah = cost_eur*(1+duty)*rate*item["qty"]
            revenue = item["price_uah"]*item["qty"]
            profit = (item["price_uah"]-cost_eur*rate)*item["qty"] if item.get("is_stock") else revenue-cost_uah
            total_profit += profit
            total_revenue += revenue

    net = total_profit - delivery
    salary = max(0, net) * SALARY_PCT / 100

    path = build_excel(name, invoices, month)
    await update.message.reply_document(
        document=open(path,"rb"),
        filename=f"ЗП_{name}_{month}.xlsx",
        caption=(
            f"📊 *Звіт за {month}*\n\n"
            f"Виторг: *{total_revenue:,.0f} грн*\n"
            f"Прибуток: *{total_profit:,.0f} грн*\n"
            f"Доставка: *{delivery:,.0f} грн*\n"
            f"Чистий прибуток: *{net:,.0f} грн*\n"
            f"━━━━━━━━━━━━\n"
            f"💰 *ЗП: {salary:,.0f} грн*"
        ),
        parse_mode="Markdown"
    )
    return ConversationHandler.END

async def cmd_clear(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    keyboard = [[InlineKeyboardButton("✅ Так",callback_data="clear_yes"),
                 InlineKeyboardButton("❌ Ні",callback_data="clear_no")]]
    await update.message.reply_text("⚠️ Видалити всі рахунки місяця?",
                                     reply_markup=InlineKeyboardMarkup(keyboard))

async def callback_clear(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    if query.data == "clear_yes":
        user = load_user(query.from_user.id)
        user["invoices"] = []
        save_user(query.from_user.id, user)
        await query.edit_message_text("✅ Очищено!")
    else:
        await query.edit_message_text("Скасовано.")

async def cmd_admin(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    if update.effective_user.id not in ADMIN_IDS:
        await update.message.reply_text("⛔ Немає доступу.")
        return
    lines = [f"📋 *Менеджери цього місяця:*\n"]
    for f in DATA_DIR.glob("*.json"):
        try:
            data = json.loads(f.read_text(encoding="utf-8"))
            name = data.get("name", f.stem)
            invoices = data.get("invoices",[])
            profit = 0
            for inv in invoices:
                for item in inv.get("items",[]):
                    lu = lookup_article(item["article"])
                    cost_eur = lu.get("cost_eur",0)
                    duty = lu.get("duty",0.04)
                    rate = item.get("rate",52.0)
                    cost_uah = cost_eur*(1+duty)*rate*item["qty"]
                    revenue = item["price_uah"]*item["qty"]
                    p = (item["price_uah"]-cost_eur*rate)*item["qty"] if item.get("is_stock") else revenue-cost_uah
                    profit += p
            salary = max(0,profit)*SALARY_PCT/100
            lines.append(f"👤 *{name}*: {len(invoices)} рахунків\n   Прибуток: {profit:,.0f} грн | ЗП: {salary:,.0f} грн\n")
        except: pass
    await update.message.reply_text("\n".join(lines), parse_mode="Markdown")

async def cmd_name(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text("Введи нове ім'я:")
    return WAIT_NAME

async def cmd_sheet(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    if SHEET_ID:
        await update.message.reply_text(
            f"📊 Google таблиця:\nhttps://docs.google.com/spreadsheets/d/{SHEET_ID}"
        )

# ── Main ──────────────────────────────────────────────────────────────────────
def main():
    app = Application.builder().token(BOT_TOKEN).build()
    conv = ConversationHandler(
        entry_points=[
            CommandHandler("start", start),
            MessageHandler(filters.Document.PDF, handle_pdf),
            CommandHandler("report", cmd_report),
            CommandHandler("name", cmd_name),
        ],
        states={
            WAIT_NAME:     [MessageHandler(filters.TEXT & ~filters.COMMAND, set_name)],
            WAIT_DATE:     [MessageHandler(filters.TEXT & ~filters.COMMAND, handle_date)],
            WAIT_STOCK:    [MessageHandler(filters.TEXT, handle_stock)],
            WAIT_DELIVERY: [MessageHandler(filters.TEXT, handle_delivery),
                            CommandHandler("0", handle_delivery)],
        },
        fallbacks=[CommandHandler("start", start)],
        allow_reentry=True,
    )
    app.add_handler(conv)
    app.add_handler(CallbackQueryHandler(callback_stock, pattern="^stock_"))
    app.add_handler(CommandHandler("clear", cmd_clear))
    app.add_handler(CommandHandler("admin", cmd_admin))
    app.add_handler(CommandHandler("sheet", cmd_sheet))
    app.add_handler(CallbackQueryHandler(callback_clear, pattern="^clear_"))
    logger.info("Bot started")
    app.run_polling(drop_pending_updates=True)

if __name__ == "__main__":
    main()
