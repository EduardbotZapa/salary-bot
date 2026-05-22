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

# Price list (prais EUR) for stock items
try:
    with open("price_lookup.json", encoding="utf-8") as f:
        PRICE_LOOKUP: dict = json.load(f)
except:
    PRICE_LOOKUP: dict = {}

def get_price(art: str) -> float:
    return PRICE_LOOKUP.get(art, 0)

def shorten_company(name: str) -> str:
    """Convert full legal name to short abbreviation + name"""
    if not name:
        return name
    name = name.strip()
    
    replacements = [
        # Full forms -> abbreviations
        ('ТОВАРИСТВО З ОБМЕЖЕНОЮ ВІДПОВІДАЛЬНІСТЮ', 'ТОВ'),
        ('Товариство з обмеженою відповідальністю', 'ТОВ'),
        ('АКЦІОНЕРНЕ ТОВАРИСТВО', 'АТ'),
        ('Акціонерне товариство', 'АТ'),
        ('ПУБЛІЧНЕ АКЦІОНЕРНЕ ТОВАРИСТВО', 'ПАТ'),
        ('ПРИВАТНЕ АКЦІОНЕРНЕ ТОВАРИСТВО', 'ПРАТ'),
        ('ФІЗИЧНА ОСОБА ПІДПРИЄМЕЦЬ', 'ФОП'),
        ('Фізична особа підприємець', 'ФОП'),
        ('ФІЗИЧНА ОСОБА-ПІДПРИЄМЕЦЬ', 'ФОП'),
        ('Фізична особа-підприємець', 'ФОП'),
        ('ДЕРЖАВНЕ ПІДПРИЄМСТВО', 'ДП'),
        ('КОМУНАЛЬНЕ ПІДПРИЄМСТВО', 'КП'),
        ('ПРИВАТНЕ ПІДПРИЄМСТВО', 'ПП'),
        ('Приватне підприємство', 'ПП'),
    ]
    
    for full, short in replacements:
        if full in name:
            name = name.replace(full, short).strip()
            break
    
    # Remove quotes around company name
    import re as re_co
    name = re_co.sub(r'[\u0022\u201c\u201d\u00ab\u00bb\u2018\u2019]', '', name).strip()
    # Remove extra spaces
    name = ' '.join(name.split())
    
    return name

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
                   "Кть","Закуп EUR","Мито%","Курс",
                   "Собів UAH/шт","Собів загал",
                   "Ціна прод UAH","Виторг",
                   "Прибуток (S)","Надбавка (T)","Склад?",
                   "УКТЗЕД","Бренд","Джерело",
                   "Прайс EUR","Вага/шт","Вага Китай","Вага Європа","Додано"]
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

        rate = inv.get("rate", 52.0)
        client = inv.get("client", "")
        invoice_num = inv.get("invoice_num", "")
        date = inv.get("date", "")
        items = inv.get("items", [])

        # ── Build all rows to write at once ──────────────────────────────────
        # Step 1: count current rows to know starting index
        existing = ws_mgr.get_all_values()
        start_row = len(existing) + 1

        all_rows_data = []   # plain values for batch append
        format_tasks = []    # (range, format_dict)
        formula_updates = [] # (range, [[formula]])

        cur_row = start_row

        # Separator row (if not first invoice)
        if len(existing) > 1:
            all_rows_data.append([""] * 24)
            format_tasks.append((f"A{cur_row}:X{cur_row}", {
                "backgroundColor": {"red": 0.85, "green": 0.91, "blue": 0.97}
            }))
            cur_row += 1

        # Invoice header row
        title = [manager_name, client, invoice_num, date] + [""] * 20
        all_rows_data.append(title)
        format_tasks.append((f"A{cur_row}:X{cur_row}", {
            "backgroundColor": {"red": 0.78, "green": 0.87, "blue": 0.95},
            "textFormat": {"bold": True}
        }))
        cur_row += 1

        # Item rows
        item_rows = []
        for item in items:
            lu = lookup_article(item["article"])
            cost_eur = lu.get("cost_eur", 0)
            duty = lu.get("duty", 0.04)
            price_eur = get_price(item["article"])
            weight_unit = lu.get("weight", 0)
            source = lu.get("source", "")
            is_stock = item.get("is_stock", False)
            duty_pct = round(duty * 100, 1)
            r = cur_row

            # Plain row first (for batch insert)
            plain_row = [
                manager_name, client, invoice_num, date,
                item["article"], item["qty"],
                round(cost_eur, 2), duty_pct, round(rate, 2),
                "", "",  # J, K - will be formula
                item["price_uah"], "",  # L, M
                "", "",  # N, O
                "так" if is_stock else "",
                lu.get("uktved", ""), lu.get("brand", ""), source,
                round(price_eur, 2) if price_eur else "",
                round(weight_unit, 3) if weight_unit else "",
                "", "",  # V, W
                datetime.now().strftime("%d.%m.%Y %H:%M"),
            ]
            all_rows_data.append(plain_row)

            # Formulas to apply after insert
            if is_stock:
                fn = f"=M{r}-T{r}*I{r}*F{r}"
                fo = f"=(M{r}-K{r})-N{r}"
            else:
                fn = f"=M{r}-K{r}"
                fo = "=0"

            formula_updates.append((f"J{r}", f"=G{r}*(1+H{r}/100)*I{r}"))
            formula_updates.append((f"K{r}", f"=J{r}*F{r}"))
            formula_updates.append((f"M{r}", f"=L{r}*F{r}"))
            formula_updates.append((f"N{r}", fn))
            formula_updates.append((f"O{r}", fo))
            formula_updates.append((f"V{r}", f'=IF(P{r}="так",IF(REGEXMATCH(LOWER(S{r}),"китай"),U{r}*F{r},0),0)'))
            formula_updates.append((f"W{r}", f'=IF(P{r}="так",IF(REGEXMATCH(LOWER(S{r}),"e-trade"),U{r}*F{r},0),0)'))

            if is_stock:
                format_tasks.append((f"A{r}:X{r}", {
                    "backgroundColor": {"red": 0.99, "green": 0.87, "blue": 0.87}
                }))

            item_rows.append((r, plain_row, is_stock))
            cur_row += 1

        # ── Step 2: Batch insert all rows at once ─────────────────────────────
        ws_mgr.append_rows(all_rows_data, value_input_option="USER_ENTERED")

        # ── Step 3: Apply formulas ────────────────────────────────────────────
        for cell_addr, formula in formula_updates:
            ws_mgr.update(cell_addr, formula, value_input_option="USER_ENTERED")

        # ── Step 4: Apply formatting ──────────────────────────────────────────
        for range_str, fmt in format_tasks:
            ws_mgr.format(range_str, fmt)

        # ── Step 5: Write to ВСІ sheet (values only, fast) ───────────────────
        all_plain = []
        for row_data in all_rows_data:
            all_plain.append(row_data)
        ws_all.append_rows(all_plain, value_input_option="USER_ENTERED")

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

# ── Minfin interbank rate ─────────────────────────────────────────────────────
async def get_nbu_rate(date_str: str):
    """Get EUR interbank buy rate from minfin.com.ua + markup %"""
    try:
        dt = datetime.strptime(date_str, "%d.%m.%Y")
        # minfin archive URL format: /currency/mb/eur/YYYY-MM-DD/
        url = f"https://minfin.com.ua/currency/mb/eur/{dt.day:02d}-{dt.month:02d}-{dt.year}/"
        headers = {
            "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36",
            "Accept": "text/html,application/xhtml+xml",
            "Accept-Language": "uk,ru;q=0.9",
        }
        async with httpx.AsyncClient(timeout=10, follow_redirects=True) as client:
            r = await client.get(url, headers=headers)
            text = r.text

        # Find buy rate in page - pattern: "можна було купити ... по курсу межбанка X"
        import re as re_mod
        # Try JSON data embedded in page
        m = re_mod.search(r'"buy"\s*:\s*"?([\d.]+)"?', text)
        if not m:
            # Try alternative pattern from page text
            m = re_mod.search(r'по курсу межбанка\s+([\d.,]+)', text)
        if not m:
            # Try to find rate in table data
            m = re_mod.search(r"[0-9]{2}[.,][0-9]{3,4}", text)

        if m:
            rate_str = m.group(1).replace(',', '.')
            rate = float(rate_str)
            if 30 < rate < 200:  # sanity check
                final = round(rate * (1 + RATE_MARKUP / 100), 2)
                logger.info(f"Minfin rate {date_str}: {rate} + {RATE_MARKUP}% = {final}")
                return final

        # Fallback: try minfin API endpoint
        api_url = f"https://minfin.com.ua/api/currency/mb/?currency=eur&date={dt:%Y-%m-%d}"
        async with httpx.AsyncClient(timeout=8) as client:
            r2 = await client.get(api_url, headers=headers)
            if r2.status_code == 200:
                data = r2.json()
                if isinstance(data, list) and data:
                    rate = float(data[0].get("buy", 0) or data[0].get("rate", 0))
                    if rate > 0:
                        return round(rate * (1 + RATE_MARKUP / 100), 2)
                elif isinstance(data, dict):
                    rate = float(data.get("buy", 0) or data.get("rate", 0))
                    if rate > 0:
                        return round(rate * (1 + RATE_MARKUP / 100), 2)

    except Exception as e:
        logger.warning(f"Minfin rate error: {e}")

    # Final fallback: NBU
    try:
        dt = datetime.strptime(date_str, "%d.%m.%Y")
        url = f"https://bank.gov.ua/NBU_Exchange/exchange_site?start={dt:%Y%m%d}&end={dt:%Y%m%d}&valcode=EUR&sort=exchangedate&order=desc&json"
        async with httpx.AsyncClient(timeout=8) as client:
            r = await client.get(url)
            data = r.json()
            if data:
                rate = float(data[0]["rate"])
                logger.info(f"Fallback NBU rate: {rate}")
                return round(rate * (1 + RATE_MARKUP / 100), 2)
    except Exception as e:
        logger.warning(f"NBU fallback error: {e}")

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
        result["client"] = shorten_company(m2.group(1).strip()[:120])

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
               'Курс','Собів UAH','Ціна UAH','Виторг','Собів загал',
               'Прибуток S','Надбавка T','Склад?','Прайс EUR','Вага/шт','Вага Китай','Вага Європа']
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
            profit = (item["price_uah"]-cost_eur*(1+duty)*rate)*item["qty"] if item.get("is_stock") else revenue-cost_total
            is_stock = item.get("is_stock",False)
            price_eur = get_price(item["article"])
            w_unit = lu.get("weight",0)
            source = lu.get("source","")
            wx = round(revenue - price_eur*rate*item["qty"],2) if (is_stock and price_eur) else 0
            s_val = wx if is_stock else round(profit,2)
            t_val = round((revenue-cost_total)-wx,2) if is_stock else 0
            w_china = round(w_unit*item["qty"],3) if (is_stock and "китай" in source.lower()) else 0
            w_eu    = round(w_unit*item["qty"],3) if (is_stock and "e-trade" in source.lower()) else 0
            vals = [inv.get("client",""),inv.get("invoice_num",""),inv.get("date",""),
                    item["article"],item["qty"],cost_eur,f'{duty*100:.0f}%',rate,
                    round(cost_unit,2),item["price_uah"],round(revenue,2),round(cost_total,2),
                    s_val,t_val,"так" if is_stock else "",price_eur,round(w_unit,3),w_china,w_eu]
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
WAIT_NAME, WAIT_DATE, WAIT_STOCK, WAIT_DELIVERY, WAIT_EXCEL = range(5)

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
            + ("\n/admin — всі менеджери\n/update — оновити довідник" if update.effective_user.id in ADMIN_IDS else "")
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
    if not_found:
        lines.append(f"⚠️ Не знайдено ({len(not_found)}): {', '.join(not_found)}")

    # Auto-get date from lookup (Дата підтвердження замовлення)
    # Use date from first found item, or today
    auto_date = ""
    for item in parsed["items"]:
        lu = lookup_article(item["article"])
        d = lu.get("confirm_date", "")
        if d:
            auto_date = d
            break

    if not auto_date:
        auto_date = datetime.now().strftime("%d.%m.%Y")

    ctx.user_data["pending_invoice"]["date"] = auto_date

    # Auto-fetch rate for that date
    rate = await get_nbu_rate(auto_date)
    if rate:
        ctx.user_data["pending_invoice"]["rate"] = rate
        for item in ctx.user_data["pending_invoice"]["items"]:
            item["rate"] = rate
        lines.append(f"\n💱 Дата: *{auto_date}* | Курс: *{rate:.2f} грн/EUR* (+{RATE_MARKUP}%)")
    else:
        ctx.user_data["pending_invoice"]["rate"] = 52.0
        lines.append(f"\n💱 Дата: *{auto_date}* | ⚠️ Курс не знайдено, використовую 52.00")

    # Show stock keyboard
    items = ctx.user_data["pending_invoice"].get("items", [])
    ctx.user_data["stock_selected"] = set()
    await msg.edit_text("\n".join(lines), parse_mode="Markdown")
    await msg.reply_text(
        "📦 *Вибери складські товари* (натисни щоб відмітити):",
        reply_markup=build_stock_keyboard(items, set()),
        parse_mode="Markdown"
    )
    return WAIT_STOCK

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
        rate_msg = f"💱 Курс міжбанк (Мінфін) на {date_str}: *{rate:.2f} грн/EUR* (+{RATE_MARKUP}%)"
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
        profit = (item["price_uah"] - cost_eur * (1 + duty) * rate) * item["qty"] if item.get("is_stock") else revenue - cost_uah
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
            profit = (item["price_uah"]-cost_eur*(1+duty)*rate)*item["qty"] if item.get("is_stock") else revenue-cost_uah
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
                    p = (item["price_uah"]-cost_eur*(1+duty)*rate)*item["qty"] if item.get("is_stock") else revenue-cost_uah
                    profit += p
            salary = max(0,profit)*SALARY_PCT/100
            lines.append(f"👤 *{name}*: {len(invoices)} рахунків\n   Прибуток: {profit:,.0f} грн | ЗП: {salary:,.0f} грн\n")
        except: pass
    await update.message.reply_text("\n".join(lines), parse_mode="Markdown")

async def cmd_name(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text("Введи нове ім'я:")
    return WAIT_NAME

async def cmd_update(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    """Admin only: update lookup from Excel file"""
    if update.effective_user.id not in ADMIN_IDS:
        await update.message.reply_text("⛔ Немає доступу.")
        return
    await update.message.reply_text(
        "📂 Надішли Excel файл таблиці замовлень (.xlsx)\n"
        "Аркуші мають називатись: *E-Trade Automation* та *Китай*",
        parse_mode="Markdown"
    )
    return WAIT_EXCEL

async def handle_excel_update(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    """Process uploaded Excel and rebuild lookup.json"""
    if update.effective_user.id not in ADMIN_IDS:
        await update.message.reply_text("⛔ Немає доступу.")
        return ConversationHandler.END

    doc = update.message.document
    if not doc.file_name.endswith('.xlsx'):
        await update.message.reply_text("❌ Потрібен .xlsx файл")
        return WAIT_EXCEL

    msg = await update.message.reply_text("⏳ Оновлюю довідник...")

    try:
        import pandas as pd

        file = await ctx.bot.get_file(doc.file_id)
        xlsx_path = str(DATA_DIR / f"orders_{doc.file_id}.xlsx")
        await file.download_to_drive(xlsx_path)

        xl = pd.ExcelFile(xlsx_path)
        sheets = xl.sheet_names
        
        dfs = []
        for sheet in sheets:
            try:
                df = pd.read_excel(xl, sheet_name=sheet, header=0)
                df['_src'] = sheet
                dfs.append(df)
            except Exception as e:
                logger.warning(f"Sheet {sheet} error: {e}")

        if not dfs:
            await msg.edit_text("❌ Не вдалось прочитати аркуші")
            return ConversationHandler.END

        import os as os_mod
        os_mod.remove(xlsx_path)

        # Find article and price columns automatically
        lookup = {}
        total = 0
        for df in dfs:
            src = df['_src'].iloc[0] if '_src' in df.columns else ''
            cols = [str(c).strip().lower() for c in df.columns]
            
            def find_col(keywords):
                for kw in keywords:
                    for i, c in enumerate(cols):
                        if kw in c: return i
                return -1

            art_i     = find_col(['артикул'])
            price_i   = find_col(['ціна за одиницю', 'price'])
            uktved_i  = find_col(['код товару', 'уктзед'])
            duty_i    = find_col(['мито'])
            weight_i  = find_col(['нетто за 1', 'вага', 'weight'])
            brand_i   = find_col(['виробник', 'brand'])

            if art_i < 0:
                continue

            for _, row in df.iterrows():
                art = str(row.iloc[art_i]).strip()
                if not art or art == 'nan' or art == 'Артикул':
                    continue
                try:
                    price = float(str(row.iloc[price_i]).replace(',','.').replace(' ','')) if price_i >= 0 else 0
                except: price = 0
                try:
                    duty = float(str(row.iloc[duty_i]).replace(',','.').replace(' ','').replace('%','')) if duty_i >= 0 else 0.04
                    if duty > 1: duty = duty / 100
                except: duty = 0.04
                try:
                    weight = float(str(row.iloc[weight_i]).replace(',','.').replace(' ','')) if weight_i >= 0 else 0
                except: weight = 0
                try:
                    brand = str(row.iloc[brand_i]).strip() if brand_i >= 0 else ''
                except: brand = ''

                lookup[art] = {
                    'uktved': str(row.iloc[uktved_i]).strip() if uktved_i >= 0 and uktved_i < len(row) else '',
                    'duty': duty,
                    'cost_eur': price,
                    'weight': weight,
                    'brand': brand if brand != 'nan' else '',
                    'source': src,
                }
                total += 1

        # Save new lookup.json
        import json as json_mod
        with open('lookup.json', 'w', encoding='utf-8') as f:
            json_mod.dump(lookup, f, ensure_ascii=False, indent=2)

        # Also update in-memory LOOKUP
        LOOKUP.clear()
        LOOKUP.update(lookup)

        await msg.edit_text(
            f"✅ *Довідник оновлено!*\n\n"
            f"📦 Артикулів: *{len(lookup)}*\n"
            f"📋 Аркушів: {len(dfs)} ({', '.join(sheets[:3])})",
            parse_mode="Markdown"
        )

    except Exception as e:
        logger.error(f"Excel update error: {e}")
        await msg.edit_text(f"❌ Помилка: {e}")

    return ConversationHandler.END

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
            CommandHandler("update", cmd_update),
        ],
        states={
            WAIT_NAME:     [MessageHandler(filters.TEXT & ~filters.COMMAND, set_name)],
            WAIT_STOCK:    [
                CallbackQueryHandler(callback_stock, pattern="^stock_"),
                MessageHandler(filters.TEXT & ~filters.COMMAND, handle_stock),
            ],
            WAIT_DELIVERY: [MessageHandler(filters.TEXT, handle_delivery),
                            CommandHandler("0", handle_delivery)],
            WAIT_EXCEL: [MessageHandler(filters.Document.MimeType(
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            ), handle_excel_update)],
        },
        fallbacks=[CommandHandler("start", start)],
        allow_reentry=True,
    )
    app.add_handler(conv)
    app.add_handler(CommandHandler("clear", cmd_clear))
    app.add_handler(CommandHandler("update", cmd_update))
    app.add_handler(CommandHandler("admin", cmd_admin))
    app.add_handler(CommandHandler("sheet", cmd_sheet))
    app.add_handler(CallbackQueryHandler(callback_clear, pattern="^clear_"))
    logger.info("Bot started")
    app.run_polling(drop_pending_updates=True)

if __name__ == "__main__":
    main()
