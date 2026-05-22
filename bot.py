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

# Invoice-specific lookup: "invoice_num:article" -> record
try:
    with open("invoice_lookup.json", encoding="utf-8") as f:
        INVOICE_LOOKUP: dict = json.load(f)
except:
    INVOICE_LOOKUP: dict = {}

# Currency rates archive: "DD.MM.YYYY" -> buy rate
try:
    with open("rates.json", encoding="utf-8") as f:
        RATES: dict = json.load(f)
except:
    RATES: dict = {}

def get_rate_from_archive(date_str: str) -> float:
    """Get EUR buy rate from local archive, return 0 if not found"""
    return float(RATES.get(date_str, 0))

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
    """Write invoice to Google Sheets using BATCH updates for speed."""
    try:
        sh = get_gsheet()
        if not sh:
            return False
        ws_mgr = get_or_create_ws(sh, manager_name)
        ws_all = get_or_create_ws(sh, "ВСІ")

        rate = inv.get("rate", 0)
        client = inv.get("client", "")
        invoice_num = inv.get("invoice_num", "")
        date = inv.get("date", "")
        items = inv.get("items", [])
        import re as re_sh
        m_sh = re_sh.search(r"[№#No]+\s*(\d+)", invoice_num)
        inv_number = m_sh.group(1) if m_sh else ""

        # Get current row count
        existing = ws_mgr.get_all_values()
        start_row = len(existing) + 1

        # Build ALL rows in memory first
        rows_to_write = []
        format_requests = []
        all_sheet_rows = []  # for ВСІ

        # Separator row (if not first entry)
        if start_row > 2:
            rows_to_write.append([""] * 24)
            format_requests.append({
                "range": f"A{start_row}:X{start_row}",
                "format": {"backgroundColor": {"red": 0.85, "green": 0.91, "blue": 0.97}}
            })
            all_sheet_rows.append([""] * 24)
            start_row += 1

        # Invoice header row
        header = [manager_name, client, invoice_num, date] + [""] * 20
        rows_to_write.append(header)
        format_requests.append({
            "range": f"A{start_row}:X{start_row}",
            "format": {
                "backgroundColor": {"red": 0.78, "green": 0.87, "blue": 0.95},
                "textFormat": {"bold": True}
            }
        })
        all_sheet_rows.append(header)
        start_row += 1

        # Item rows with formulas
        stock_rows = []  # rows to color red on E column
        for item in items:
            lu = lookup_article(item["article"], inv_number)
            cost_eur = lu.get("cost_eur", 0)
            duty = lu.get("duty", 0.04)
            price_eur = get_price(item["article"])
            weight_unit = lu.get("weight", 0)
            source = lu.get("source", "")
            is_stock = item.get("is_stock", False)
            duty_pct = round(duty * 100, 1)
            r = start_row

            if is_stock:
                fn = f"=M{r}-T{r}*I{r}*F{r}"
                fo = f"=(M{r}-K{r})-N{r}"
            else:
                fn = f"=M{r}-K{r}"
                fo = "=0"

            row = [
                "",                                              # A Менеджер
                "",                                              # B Клієнт
                "",                                              # C Рахунок
                "",                                              # D Дата
                item["article"],                                 # E Артикул
                item["qty"],                                     # F Кть
                round(cost_eur, 2),                              # G Закуп EUR
                duty_pct,                                        # H Мито%
                round(rate, 2) if rate else "",                  # I Курс
                f"=G{r}*(1+H{r}/100)*I{r}",                      # J Собів UAH/шт
                f"=J{r}*F{r}",                                   # K Собів загал
                item["price_uah"],                               # L Ціна прод
                f"=L{r}*F{r}",                                   # M Виторг
                fn,                                              # N Прибуток S
                fo,                                              # O Надбавка T
                "так" if is_stock else "",                       # P Склад?
                lu.get("uktved", ""),                            # Q УКТЗЕД
                lu.get("brand", ""),                             # R Бренд
                source,                                          # S Джерело
                round(price_eur, 2) if price_eur else "",        # T Прайс EUR
                round(weight_unit, 3) if weight_unit else "",    # U Вага/шт
                f'=IF(P{r}="так",IF(REGEXMATCH(LOWER(S{r}),"китай"),U{r}*F{r},0),0)',  # V
                f'=IF(P{r}="так",IF(REGEXMATCH(LOWER(S{r}),"e-trade"),U{r}*F{r},0),0)', # W
                datetime.now().strftime("%d.%m.%Y %H:%M"),       # X
            ]
            rows_to_write.append(row)
            all_sheet_rows.append(row)

            if is_stock:
                stock_rows.append(r)

            start_row += 1

        # ── ONE batch write to manager sheet ──────────────────────────────────
        first_row = len(existing) + 1
        last_row = first_row + len(rows_to_write) - 1
        ws_mgr.update(
            f"A{first_row}:X{last_row}",
            rows_to_write,
            value_input_option="USER_ENTERED"
        )

        # ── Number formatting (batch) ─────────────────────────────────────────
        num_format = {"numberFormat": {"type": "NUMBER", "pattern": "#,##0.00"}}
        # Only data rows (skip separator and header)
        data_start = first_row + (2 if existing and len(existing) > 1 else 1)
        if data_start <= last_row:
            for col in ["G", "H", "I", "J", "K", "L", "M", "N", "O", "T", "U", "V", "W"]:
                ws_mgr.format(f"{col}{data_start}:{col}{last_row}", num_format)

        # ── Color stock article cells ─────────────────────────────────────────
        for sr in stock_rows:
            ws_mgr.format(f"E{sr}", {
                "backgroundColor": {"red": 0.99, "green": 0.87, "blue": 0.87},
                "textFormat": {"bold": True, "foregroundColor": {"red": 0.8, "green": 0.0, "blue": 0.0}}
            })

        # ── Apply separator/header formatting ─────────────────────────────────
        for req in format_requests:
            ws_mgr.format(req["range"], req["format"])

        # ── Append to ВСІ sheet (single batch) ────────────────────────────────
        try:
            ws_all.append_rows(all_sheet_rows, value_input_option="USER_ENTERED")
        except Exception as e:
            logger.warning(f"ВСІ append error: {e}")

        return True
    except Exception as e:
        logger.error(f"Sheet append error: {e}")
        import traceback
        logger.error(traceback.format_exc())
        return False

# ── Storage ───────────────────────────────────────────────────────────────────
def _uf(uid): return DATA_DIR / f"{uid}.json"
def load_user(uid):
    p = _uf(uid)
    return json.loads(p.read_text(encoding="utf-8")) if p.exists() else {"name":"","invoices":[]}
def save_user(uid, data):
    _uf(uid).write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")

# ── Currency rate (archive first, then Minfin, then NBU) ─────────────────────
async def get_nbu_rate(date_str: str):
    """Get EUR buy rate: archive -> Minfin -> NBU fallback"""
    import re as re_rate

    # ── Method 0: Local rates archive (most accurate) ─────────────────────────
    archived = get_rate_from_archive(date_str)
    if archived > 0:
        final = round(archived * (1 + RATE_MARKUP / 100), 2)
        logger.info(f"Archive rate {date_str}: {archived} -> {final}")
        return final

    try:
        dt = datetime.strptime(date_str, "%d.%m.%Y")

        headers = {
            "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36",
            "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
            "Accept-Language": "uk-UA,uk;q=0.9,ru;q=0.8",
            "Referer": "https://minfin.com.ua/currency/mb/",
        }

        # ── Method 1: Minfin EUR archive page ────────────────────────────────
        # URL: /currency/mb/eur/DD-MM-YYYY/
        page_url = f"https://minfin.com.ua/currency/mb/eur/{dt.day:02d}-{dt.month:02d}-{dt.year}/"
        async with httpx.AsyncClient(timeout=12, follow_redirects=True) as client:
            r = await client.get(page_url, headers=headers)
            text = r.text

        # Look for: "по курсу межбанка 51,4287"
        m = re_rate.search(r"по курсу межбанка\s+([\d]+[,.][\d]+)", text)
        if m:
            rate = float(m.group(1).replace(",", "."))
            if 48 < rate < 65:  # EUR/UAH realistic range
                final = round(rate * (1 + RATE_MARKUP / 100), 2)
                logger.info(f"Minfin EUR buy {date_str}: {rate} -> {final}")
                return final

        # Look for JSON in __NEXT_DATA__ or similar script tags
        # Pattern: "buy":"51.4287" near "eur" context
        eur_section = text.lower()
        eur_pos = eur_section.find('"eur"')
        if eur_pos == -1:
            eur_pos = eur_section.find("euro")
        if eur_pos > 0:
            chunk = text[max(0, eur_pos-200):eur_pos+500]
            m = re_rate.search(r'"buy"\s*:\s*"?([\d.]+)"?', chunk)
            if m:
                rate = float(m.group(1))
                if 48 < rate < 65:
                    final = round(rate * (1 + RATE_MARKUP / 100), 2)
                    logger.info(f"Minfin JSON EUR buy {date_str}: {rate} -> {final}")
                    return final

        # ── Method 2: search all "buy" values, pick EUR range ────────────────
        all_buys = re_rate.findall(r'"buy"\s*:\s*"?([\d.]+)"?', text)
        for b in all_buys:
            rate = float(b)
            if 48 < rate < 65:  # EUR range
                final = round(rate * (1 + RATE_MARKUP / 100), 2)
                logger.info(f"Minfin all-buy EUR {date_str}: {rate} -> {final}")
                return final

        # ── Method 3: find any number in EUR range on EUR page ───────────────
        all_nums = re_rate.findall(r"5[0-9][,.][\d]{4}", text)
        if all_nums:
            rate = float(all_nums[0].replace(",", "."))
            if 48 < rate < 65:
                final = round(rate * (1 + RATE_MARKUP / 100), 2)
                logger.info(f"Minfin num EUR {date_str}: {rate} -> {final}")
                return final

    except Exception as e:
        logger.warning(f"Minfin rate error: {e}")

    # ── Fallback: NBU official rate ───────────────────────────────────────────
    try:
        dt = datetime.strptime(date_str, "%d.%m.%Y")
        url = f"https://bank.gov.ua/NBU_Exchange/exchange_site?start={dt:%Y%m%d}&end={dt:%Y%m%d}&valcode=EUR&sort=exchangedate&order=desc&json"
        async with httpx.AsyncClient(timeout=8) as client:
            r = await client.get(url)
            data = r.json()
            if data:
                rate = float(data[0]["rate"])
                logger.info(f"NBU fallback EUR {date_str}: {rate}")
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

def lookup_article(art: str, invoice_num: str = "") -> dict:
    art = art.strip()
    # Try invoice-specific lookup first (most accurate)
    if invoice_num:
        inv_key = f"{invoice_num}:{art}"
        if inv_key in INVOICE_LOOKUP:
            rec = INVOICE_LOOKUP[inv_key]
            if rec.get("cost_eur", 0) > 0:
                return rec
    # Try live cache
    if ORDERS_SHEET_ID and _live_cache:
        result = _live_cache.get(art)
        if result:
            return result
    # Exact match in static lookup
    if art in LOOKUP:
        return LOOKUP[art]
    # Normalized match (handle extra spaces)
    art_norm = " ".join(art.upper().split())
    for key, val in LOOKUP.items():
        if " ".join(key.upper().split()) == art_norm:
            return val
    return {}

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
            lu = lookup_article(item["article"], inv_number)
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
WAIT_NAME, WAIT_DATE, WAIT_STOCK, WAIT_DELIVERY, WAIT_EXCEL, WAIT_RATES = range(6)

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
            + ("\n/admin — всі менеджери\n/update — оновити довідник\n/rates — оновити курси валют" if update.effective_user.id in ADMIN_IDS else "")
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

    # Extract invoice number for precise lookup
    inv_number = ""
    import re as re_inv
    m_inv = re_inv.search(r'[№#No]+\s*(\d+)', parsed.get("invoice_num", ""))
    if m_inv:
        inv_number = m_inv.group(1)

    found, not_found = [], []
    for item in parsed["items"]:
        lu = lookup_article(item["article"], inv_number)
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

    # ── Date & rate logic ────────────────────────────────────────────────────
    # Try to find date from non-stock item in orders table
    auto_date = ""
    for item in parsed["items"]:
        lu = lookup_article(item["article"], inv_number)
        d = lu.get("confirm_date", "")
        if d:
            auto_date = d
            break

    items = ctx.user_data["pending_invoice"].get("items", [])
    ctx.user_data["stock_selected"] = set()

    if auto_date:
        # Found date from orders table - get rate automatically
        rate = await get_nbu_rate(auto_date)
        if rate:
            ctx.user_data["pending_invoice"]["rate"] = rate
            ctx.user_data["pending_invoice"]["date"] = auto_date
            for item in ctx.user_data["pending_invoice"]["items"]:
                item["rate"] = rate
            lines.append(f"\n💱 Дата: *{auto_date}* | Курс: *{rate:.2f} грн/EUR* (+{RATE_MARKUP}%)")
        else:
            ctx.user_data["pending_invoice"]["rate"] = 0
            ctx.user_data["pending_invoice"]["date"] = auto_date
            lines.append(f"\n💱 Дата: *{auto_date}* | ⚠️ Курс не знайдено — вкажи вручну в таблиці")

        await msg.edit_text("\n".join(lines), parse_mode="Markdown")
        await msg.reply_text(
            "📦 *Вибери складські товари* (натисни щоб відмітити):",
            reply_markup=build_stock_keyboard(items, set()),
            parse_mode="Markdown"
        )
        return WAIT_STOCK
    else:
        # All items are stock or not found - ask manager for date
        ctx.user_data["pending_invoice"]["rate"] = 0
        ctx.user_data["pending_invoice"]["date"] = ""
        lines.append(f"\n📅 Всі товари зі складу — введи дату оплати клієнтом (ДД.ММ.РРРР):")
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
        rate_msg = f"💱 Курс міжбанк (Мінфін) на {date_str}: *{rate:.2f} грн/EUR* (+{RATE_MARKUP}%)"
    else:
        ctx.user_data["pending_invoice"]["rate"] = 52.0
        rate_msg = f"⚠️ Не вдалось отримати курс. Використовую 52.00"

    inv = ctx.user_data.get("pending_invoice", {})
    items = inv.get("items", [])
    ctx.user_data["stock_selected"] = set()
    await update.message.reply_text(rate_msg, parse_mode="Markdown")
    await update.message.reply_text(
        "📦 *Вибери складські товари* (натисни щоб відмітити):\n_(для складського рахунку можна одразу натиснути Зберегти)_",
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
            lu = lookup_article(item["article"], inv_number)
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
                    lu = lookup_article(item["article"], inv_number)
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

async def cmd_rates(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    """Admin only: update currency rates from Excel file"""
    if update.effective_user.id not in ADMIN_IDS:
        await update.message.reply_text("⛔ Немає доступу.")
        return
    await update.message.reply_text(
        "📂 Надішли Excel файл з курсами валют (.xlsx)\n"
        "Формат: колонка *Дата* і *Курс покупки* (або аналогічна структура)",
        parse_mode="Markdown"
    )
    return WAIT_RATES

async def handle_rates_update(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    """Process uploaded rates Excel"""
    if update.effective_user.id not in ADMIN_IDS:
        return ConversationHandler.END

    doc = update.message.document
    if not doc.file_name.endswith(".xlsx"):
        await update.message.reply_text("❌ Потрібен .xlsx файл")
        return WAIT_RATES

    msg = await update.message.reply_text("⏳ Оновлюю курси...")
    try:
        import pandas as pd

        file = await ctx.bot.get_file(doc.file_id)
        xlsx_path = str(DATA_DIR / f"rates_{doc.file_id}.xlsx")
        await file.download_to_drive(xlsx_path)

        xl = pd.ExcelFile(xlsx_path)
        df = pd.read_excel(xl, sheet_name=0, header=1)
        df.columns = [str(c).strip() for c in df.columns]
        import os as os_mod
        os_mod.remove(xlsx_path)

        # Find date and buy columns
        date_col = next((c for c in df.columns if 'дат' in c.lower() or 'date' in c.lower()), df.columns[0])
        # Accept both buy/sell column names - just take the rate column (2nd column)
        buy_col = next(
            (c for c in df.columns if 'покуп' in c.lower() or 'buy' in c.lower() or 'курс' in c.lower()),
            df.columns[1]
        )

        rates_new = {}
        for _, row in df.iterrows():
            try:
                d = row[date_col]
                if pd.isna(d): continue
                date_str = d.strftime("%d.%m.%Y") if hasattr(d, "strftime") else str(d)[:10]
                buy = float(str(row[buy_col]).replace(",", "."))
                if buy > 0:
                    rates_new[date_str] = buy
            except: pass

        with open("rates.json", "w", encoding="utf-8") as f:
            json.dump(rates_new, f, ensure_ascii=False, indent=2)

        RATES.clear()
        RATES.update(rates_new)

        await msg.edit_text(
            f"✅ *Курси оновлено!*\n\n"
            f"📅 Дат: *{len(rates_new)}*\n"
            f"📆 Від {min(rates_new.keys())} до {max(rates_new.keys())}",
            parse_mode="Markdown"
        )
    except Exception as e:
        await msg.edit_text(f"❌ Помилка: {e}")

    return ConversationHandler.END

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
            CommandHandler("rates", cmd_rates),
        ],
        states={
            WAIT_NAME:     [MessageHandler(filters.TEXT & ~filters.COMMAND, set_name)],
            WAIT_DATE:     [MessageHandler(filters.TEXT & ~filters.COMMAND, handle_date)],
            WAIT_STOCK:    [
                CallbackQueryHandler(callback_stock, pattern="^stock_"),
                MessageHandler(filters.TEXT & ~filters.COMMAND, handle_stock),
            ],
            WAIT_DELIVERY: [MessageHandler(filters.TEXT, handle_delivery),
                            CommandHandler("0", handle_delivery)],
            WAIT_RATES: [MessageHandler(filters.Document.MimeType(
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            ), handle_rates_update)],
            WAIT_EXCEL: [MessageHandler(filters.Document.MimeType(
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            ), handle_excel_update)],
        },
        fallbacks=[CommandHandler("start", start)],
        allow_reentry=True,
    )
    app.add_handler(conv)
    app.add_handler(CommandHandler("clear", cmd_clear))
    app.add_handler(CommandHandler("rates", cmd_rates))
    app.add_handler(CommandHandler("update", cmd_update))
    app.add_handler(CommandHandler("admin", cmd_admin))
    app.add_handler(CommandHandler("sheet", cmd_sheet))
    app.add_handler(CallbackQueryHandler(callback_clear, pattern="^clear_"))
    logger.info("Bot started")
    app.run_polling(drop_pending_updates=True)

if __name__ == "__main__":
    main()
