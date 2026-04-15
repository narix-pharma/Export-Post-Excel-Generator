# -*- coding: utf-8 -*-
import sys, io
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')
sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding='utf-8', errors='replace')
from flask import Flask, request, jsonify, send_file, send_from_directory
import pdfplumber
import re
from datetime import datetime
from dateutil.relativedelta import relativedelta
import openpyxl
from openpyxl import load_workbook
from openpyxl.styles import Font, Alignment, PatternFill, GradientFill, Border, Side
import io
import os

app = Flask(__name__, static_folder='static')
app.config['MAX_CONTENT_LENGTH'] = 50 * 1024 * 1024  # 50MB

PRODUCT_DATA_PATH = os.path.join(os.path.dirname(__file__), 'PRODUCT_DATA.xlsx')

def load_product_db():
    """Read PRODUCT_DATA.xlsx and return a dict keyed by stripped brand name."""
    if not os.path.exists(PRODUCT_DATA_PATH):
        return {}
    try:
        wb = load_workbook(PRODUCT_DATA_PATH, read_only=True, data_only=True)
        ws = wb.active
        db = {}
        for row in ws.iter_rows(min_row=2, values_only=True):
            if not row or not row[0]: continue
            brand = str(row[0]).strip().upper()
            if not brand or brand in ('-BLANK-','BLANK','N/A'): continue
            generic  = str(row[1]).strip().upper() if row[1] and str(row[1]).strip().upper() not in ('-BLANK-','BLANK','N/A') else ''
            mfr      = str(row[2]).strip().upper() if len(row)>2 and row[2] and str(row[2]).strip().upper() not in ('-BLANK-','BLANK','N/A') else ''
            hsn      = str(row[3]).strip().upper() if len(row)>3 and row[3] and str(row[3]).strip().upper() not in ('-BLANK-','BLANK','N/A') else ''
            key = re.sub(r'\s+\d+(?:X\d+)?(?:TAB|CAP)$', '', brand, flags=re.IGNORECASE).strip()
            key = re.sub(r'\s+\d+X\d+GM$', '', key, flags=re.IGNORECASE).strip()
            key = re.sub(r'\s+\d+\*\d+$', '', key, flags=re.IGNORECASE).strip()
            if key in db:
                if not db[key]['g'] and generic: db[key]['g'] = generic
                if not db[key]['m'] and mfr:     db[key]['m'] = mfr
                if not db[key]['h'] and hsn:      db[key]['h'] = hsn
            else:
                db[key] = {'g': generic, 'm': mfr, 'h': hsn}
        wb.close()
        print(f"  Product DB loaded: {len(db)} products from PRODUCT_DATA.xlsx")
        return db
    except Exception as e:
        print(f"  [WARN] Could not load PRODUCT_DATA.xlsx: {e}")
        return {}

PRODUCT_DB = load_product_db()
TEMPLATE_PATH = os.path.join(os.path.dirname(__file__), 'template.xlsx')

CAMBRIA_12_BOLD = Font(name='Cambria', size=12, bold=True)
CAMBRIA_14_BOLD = Font(name='Cambria', size=14, bold=True)
CAMBRIA_11      = Font(name='Cambria', size=11)
CAMBRIA_11_BOLD = Font(name='Cambria', size=11, bold=True)
CAMBRIA_10      = Font(name='Cambria', size=10)
CALIBRI_12      = Font(name='Calibri', size=12)

FMT_USD  = '"$"#,##0.00'
# Border styles for INVOICE item rows (matching format_invoice.xlsx)
THIN  = Side(border_style='thin')
MED   = Side(border_style='medium')
# Pre-built borders per column (matching format_invoice layout exactly)
# A: top+bottom only   B: all thin   C: right+top+bottom   D: L+R+T+B
# E: L+T+B (no right sometimes)   F: L+R+T+B   G: L+R+T+B
# H: L+R+T+B   I: L+R+T+B   J: L+T+B (no right)   K: L+T+B (no right)   L: L+R+T+B
def item_border(col):
    """Return the correct Border for a given column (1-indexed) in an item row."""
    T, B, L, R = THIN, THIN, THIN, THIN
    if col == 1:   return Border(top=T, bottom=B)                         # A
    if col == 2:   return Border(left=L, right=R, top=T, bottom=B)        # B
    if col == 3:   return Border(right=R, top=T, bottom=B)                # C (no left)
    if col == 4:   return Border(left=L, right=R, top=T, bottom=B)        # D
    if col == 5:   return Border(left=L, top=T, bottom=B)                 # E (no right — varies)
    if col == 6:   return Border(left=L, right=R, top=T, bottom=B)        # F
    if col == 7:   return Border(left=L, right=R, top=T, bottom=B)        # G
    if col == 8:   return Border(left=L, right=R, top=T, bottom=B)        # H
    if col == 9:   return Border(left=L, right=R, top=T, bottom=B)        # I
    if col == 10:  return Border(left=L, top=T, bottom=B)                 # J (no right)
    if col == 11:  return Border(left=L, top=T, bottom=B)                 # K (no right)
    if col == 12:  return Border(left=L, right=R, top=T, bottom=B)        # L
    return Border()

def total_row_border(col):
    """Border for the TOTAL row cells that have values."""
    T, B, L, R = THIN, THIN, THIN, THIN
    if col == 8:   return Border(left=L, right=R, top=T, bottom=B)
    if col == 9:   return Border(left=L, right=R, top=T, bottom=B)
    if col == 10:  return Border(left=L, right=R, top=T, bottom=B)
    if col == 11:  return Border(left=L, right=R, top=T, bottom=B)
    if col == 12:  return Border(left=L, right=R, top=T, bottom=B)
    return Border()

# Highlight fill for TOTAL rows in packing list (cols C to K)
TOTAL_FILL  = PatternFill(fill_type='solid', fgColor='C6EFCE')   # light green
NO_FILL     = PatternFill(fill_type=None)
FMT_INR  = '"₹"\\ #,##0.00'
FMT_DATE = 'mmm\\-yy'

def safe_rate(rate):
    """Format dollar_rate for Excel formula — never produces scientific notation."""
    from decimal import Decimal
    d = Decimal(str(float(rate)))
    return format(d.normalize(), 'f')



# ─────────────────────────────────────────────
#  PDF PARSING
# ─────────────────────────────────────────────
def parse_pdf(file_bytes):
    result = {'invoice_no': '', 'date': '', 'buyer_name': '', 'buyer_address': '', 'items': []}
    with pdfplumber.open(io.BytesIO(file_bytes)) as pdf:
        full_text = ''.join(page.extract_text() + '\n' for page in pdf.pages)

    inv_match = re.search(r'P/(\d+)\s+(\d{2}-\d{2}-\d{4})', full_text)
    if inv_match:
        result['invoice_no'] = 'P/' + inv_match.group(1)
        result['date'] = inv_match.group(2)

    # Buyer: find name after "AS PER PACKING LIST" block, then grab address
    skip_labels = {'Exporter', 'Consignee', 'NARIX PHARMACEUTICALS PVT LTD'}
    all_s3b = re.findall(r'\(s3B([^\(]+?)\(cid:27\)\(s0B', full_text)
    for candidate in all_s3b:
        c = candidate.strip()
        if c and c not in skip_labels and 'AS PER' not in c and 'PACKING' not in c:
            result['buyer_name'] = c
            break
    # Address lines immediately after buyer name in PDF table
    if result['buyer_name']:
        buyer_idx = full_text.find(result['buyer_name'])
        after = full_text[buyer_idx:]
        addr_lines = re.findall(r'\|\s*\|\s*([A-Z][^|\n]+?)\s*\|', after)
        clean = []
        for line in addr_lines:
            line = line.strip()
            if not line or 'Country' in line or line == 'INDIA' or 'AS PER' in line:
                break
            clean.append(line)
        result['buyer_address'] = '\n'.join(clean)

    lines = full_text.split('\n')
    items = []
    i = 0
    while i < len(lines):
        line = lines[i]
        item_match = re.search(
            r'\|\s+([A-Z][A-Z0-9 \-\(\)\.\/\%\+\*]+?)\s+\|\s+(\d+)\s+PCS\s+\|\s+([\d.]+)\s+\|\s+([\d.]+)\s+\|',
            line
        )
        if item_match:
            raw = item_match.group(1).strip()
            # Strip ONLY trailing pack-size suffix (10TAB, 15CAP, 14X5GM) — preserve dosage (12MG, 0.5MG)
            product = re.sub(r'\s+\d+(?:X\d+)?(?:TAB|CAP)$', '', raw, flags=re.IGNORECASE).strip().upper()
            product = re.sub(r'\s+\d+X\d+GM$', '', product, flags=re.IGNORECASE).strip()
            product = re.sub(r'\s+\d+\*\d+$', '', product, flags=re.IGNORECASE).strip()
            qty   = int(item_match.group(2))
            price = float(item_match.group(3))
            total = float(item_match.group(4))
            batch = exp_str = ''
            if i + 1 < len(lines):
                bm = re.search(r'Batch\s*:\s*(\S+)\s+Exp\.\s*:\s*(\d+/\d+)', lines[i+1])
                if bm:
                    batch, exp_str = bm.group(1), bm.group(2)
                    i += 1
            items.append({'product': product, 'qty': qty, 'price': price,
                          'total': total, 'batch': batch, 'exp': exp_str,
                          'hsn': '3004', 'generic': '', 'manufacturer': '',
                          'mfg_years_back': 2})
        i += 1

    result['items'] = items
    return result


# ─────────────────────────────────────────────
#  HELPERS
# ─────────────────────────────────────────────
def parse_exp_date(exp_str):
    try:
        m, y = exp_str.split('/')
        return datetime(int(y), int(m), 1)
    except:
        return None

def mfg_from_exp(exp_str, years_back):
    # Pharma: mfg = expiry +1 month then minus N years
    # e.g. Exp Apr/2027 2yr -> May/2027 -> May/2025
    d = parse_exp_date(exp_str)
    if not d: return None
    return d + relativedelta(months=1) - relativedelta(years=int(years_back))

def set_cell(ws, row, col, value, font=None, num_fmt=None):
    cell = ws.cell(row=row, column=col)
    cell.value = value
    if font:    cell.font = font
    if num_fmt: cell.number_format = num_fmt
    return cell


# ─────────────────────────────────────────────
#  INVOICE SHEET
# ─────────────────────────────────────────────
def fill_invoice(wb, invoice_no, inv_date_str, dollar_rate, items):
    TEMPLATE_FOOTER_START = 32
    START_ROW = 20
    n = len(items)
    total_row = START_ROW + n
    footer_start = total_row + 4

    ws = wb['INVOICE']

    # Read footer before clearing
    footer_cells = []
    for r in range(TEMPLATE_FOOTER_START, TEMPLATE_FOOTER_START + 12):
        for cell in ws[r]:
            if cell.value is not None:
                footer_cells.append({'row_offset': r - TEMPLATE_FOOTER_START,
                                     'col': cell.column, 'value': cell.value,
                                     'bold': cell.font.bold, 'size': cell.font.size})

    ws['I2'].value = invoice_no
    ws['I3'].value = inv_date_str

    for r in range(START_ROW, footer_start + 15):
        for c in range(1, 14):
            ws.cell(row=r, column=c).value = None

    for idx, item in enumerate(items):
        row     = START_ROW + idx
        exp_d   = parse_exp_date(item.get('exp', ''))
        mfg_d   = mfg_from_exp(item.get('exp', ''), item.get('mfg_years_back', 2))

        def c(col):
            cell = ws.cell(row=row, column=col)
            cell.border = item_border(col)
            return cell
        c(1).value = idx + 1;                              c(1).font = CAMBRIA_11
        c(2).value = int(item.get('hsn', 3004));           c(2).font = CAMBRIA_11
        c(3).value = item['product'].upper();               c(3).font = CAMBRIA_10
        c(4).value = item.get('generic', '').upper();       c(4).font = CAMBRIA_11
        c(5).value = item.get('batch', '');                 c(5).font = CAMBRIA_11
        c(6).value = mfg_d;                                 c(6).font = CAMBRIA_11
        c(7).value = exp_d;                                 c(7).font = CAMBRIA_11
        c(8).value = item.get('manufacturer', '').upper();  c(8).font = CAMBRIA_11
        c(9).value = item.get('qty');                       c(9).font = CAMBRIA_11
        c(10).value = item.get('price');                    c(10).font = CAMBRIA_11; c(10).number_format = FMT_USD
        c(11).value = f'=J{row}*I{row}';                   c(11).font = CAMBRIA_11; c(11).number_format = FMT_USD
        c(12).value = f'=K{row}*{safe_rate(dollar_rate)}';            c(12).font = CAMBRIA_11; c(12).number_format = FMT_INR
        if mfg_d: c(6).number_format = FMT_DATE
        if exp_d: c(7).number_format = FMT_DATE

    # Total row
    tr = total_row
    set_cell(ws, tr, 3,  'Total article Qty ---  01', CAMBRIA_14_BOLD)
    set_cell(ws, tr, 8,  'Total',                      CAMBRIA_14_BOLD)
    k_cell = set_cell(ws, tr, 11, f'=SUM(K{START_ROW}:K{tr-1})', CAMBRIA_14_BOLD, FMT_USD)
    l_cell = set_cell(ws, tr, 12, f'=SUM(L{START_ROW}:L{tr-1})', CAMBRIA_14_BOLD, FMT_INR)
    # Apply borders to total row valued cells (H through L)
    for col in range(8, 13):
        ws.cell(row=tr, column=col).border = total_row_border(col)

    # Footer
    for fc in footer_cells:
        cell = ws.cell(row=footer_start + fc['row_offset'], column=fc['col'])
        cell.value = fc['value']
        cell.font  = Font(name='Cambria', size=fc['size'], bold=fc['bold'])

    return total_row


# ─────────────────────────────────────────────
#  PACKING LIST SHEET
# ─────────────────────────────────────────────
def fill_packing_list(wb, invoice_no, inv_date_str, dollar_rate,
                      buyer_name, buyer_address, invoice_items, customers):
    """
    invoice_items: list of invoice line items (with all details)
    customers: list of {name, address, lines: [{product, qty, pills_per_strip}]}
    """
    ws = wb['PACKING LIST']

    # Build lookup: product -> invoice item details
    # Lookup by (product, batch) to handle duplicate product names
    inv_lookup = {(item['product'].upper(), item['batch']): item for item in invoice_items}
    inv_lookup_by_product = {item['product'].upper(): item for item in invoice_items}  # fallback

    # ── Header: invoice no & date (F3/H3) ───────
    ws['F3'].value = invoice_no
    ws['F3'].font  = CAMBRIA_12_BOLD
    ws['H3'].value = inv_date_str
    ws['H3'].font  = CAMBRIA_12_BOLD

    # ── Buyer details (G9 name, G10 address) ────
    ws['G9'].value  = buyer_name.upper()
    ws['G9'].font   = CAMBRIA_12_BOLD
    ws['G10'].value = buyer_address.upper()
    ws['G10'].font  = CAMBRIA_12_BOLD

    # ── Clear existing customer rows (unmerge, clear values AND fills) ──
    merges_to_remove = [str(m) for m in list(ws.merged_cells.ranges) if m.min_row >= 22]
    for m in merges_to_remove:
        ws.unmerge_cells(m)
    for r in range(22, 200):
        for c in range(1, 14):
            cell = ws.cell(row=r, column=c)
            cell.value = None
            cell.fill  = NO_FILL

    # ── Write customers ──────────────────────────
    cur_row = 22
    cust_num = 1

    for cust in customers:
        cust_name    = cust.get('name', '').upper()
        # Build display address from split fields (fallback to legacy 'address' field)
        if cust.get('addr1'):
            addr_parts = [str(cust.get('addr1','')), str(cust.get('city','')), str(cust.get('state','')), str(cust.get('zip','')), str(cust.get('country',''))]
            cust_address = ', '.join(p.upper() for p in addr_parts if p and p.strip())
        else:
            cust_address = cust.get('address', '').upper()
        full_label   = f"{cust_name}\n\n{cust_address}"
        first_item_row = cur_row

        for line_idx, line in enumerate(cust.get('lines', [])):
            product    = line['product'].upper()
            qty        = int(line['qty'])
            pills      = int(line.get('pills_per_strip', 10))
            batch_key  = (product, line.get('batch', ''))
            inv_item   = inv_lookup.get(batch_key) or inv_lookup_by_product.get(product, {})

            exp_d = parse_exp_date(inv_item.get('exp', ''))
            mfg_d = mfg_from_exp(inv_item.get('exp', ''), inv_item.get('mfg_years_back', 2))
            batch = inv_item.get('batch', '')
            price = inv_item.get('price', 0)
            mfr   = inv_item.get('manufacturer', '').upper()

            row = cur_row

            # Sr No, customer name and tracking only on first line of customer
            if line_idx == 0:
                set_cell(ws, row, 1, cust_num, CALIBRI_12)
                cell_b = ws.cell(row=row, column=2)
                cell_b.value = full_label
                cell_b.font  = CALIBRI_12
                cell_b.alignment = Alignment(wrap_text=True, vertical='top')
                tracking = cust.get('tracking', '').strip()
                if tracking:
                    set_cell(ws, row, 13, tracking, CALIBRI_12)

            set_cell(ws, row, 3, product,  CALIBRI_12)
            set_cell(ws, row, 4, batch,    CALIBRI_12)

            c_mfg = ws.cell(row=row, column=5)
            c_mfg.value = mfg_d; c_mfg.font = CALIBRI_12
            if mfg_d: c_mfg.number_format = FMT_DATE

            c_exp = ws.cell(row=row, column=6)
            c_exp.value = exp_d; c_exp.font = CALIBRI_12
            if exp_d: c_exp.number_format = FMT_DATE

            set_cell(ws, row, 7, mfr,   CALIBRI_12)
            set_cell(ws, row, 8, qty,   CALIBRI_12)
            set_cell(ws, row, 9, price, CALIBRI_12)
            set_cell(ws, row, 10, f'=H{row}*I{row}', CALIBRI_12, FMT_USD)
            set_cell(ws, row, 11, f'=J{row}*{safe_rate(dollar_rate)}', CALIBRI_12, FMT_INR)
            set_cell(ws, row, 12, f'=H{row}*{pills}', CALIBRI_12)

            cur_row += 1

        # TOTAL row for this customer
        last_item_row = cur_row - 1
        total_r = cur_row

        # Strip any leftover template fill from every row written (item rows)
        for item_r in range(first_item_row, cur_row):
            for col in range(1, 14):
                ws.cell(row=item_r, column=col).fill = NO_FILL

        # Write TOTAL cells C / J / K with highlight fill
        tc = set_cell(ws, total_r, 3,  'TOTAL', CAMBRIA_14_BOLD)
        tj = set_cell(ws, total_r, 10, f'=SUM(J{first_item_row}:J{last_item_row})', CAMBRIA_14_BOLD, FMT_USD)
        tk = set_cell(ws, total_r, 11, f'=SUM(K{first_item_row}:K{last_item_row})', CAMBRIA_14_BOLD, FMT_INR)
        # Apply fill only to TOTAL row columns C through K
        for col in range(3, 12):   # C=3 to K=11
            ws.cell(row=total_r, column=col).fill = TOTAL_FILL

        cur_row += 2   # total row + 1 blank row gap
        cust_num += 1

    # ── Signature/Co. Stamp — 5 rows after last customer TOTAL row ──
    sig_row = cur_row + 3   # cur_row is already 2 past last total; +3 = 5 rows gap total
    sig_cell = ws.cell(row=sig_row, column=11)   # column K = 11
    sig_cell.value = 'Signature/Co. Stamp'
    sig_cell.font  = CAMBRIA_12_BOLD


# ─────────────────────────────────────────────
#  MAIN EXCEL GENERATOR
# ─────────────────────────────────────────────
def generate_excel(invoice_data):
    wb          = load_workbook(TEMPLATE_PATH)
    invoice_no  = invoice_data['invoice_no']
    date_str    = invoice_data['date']
    dollar_rate = float(invoice_data.get('dollar_rate', 90.2))
    items       = invoice_data['items']
    customers   = invoice_data.get('customers', [])
    buyer_name  = invoice_data.get('buyer_name', '')
    buyer_addr  = invoice_data.get('buyer_address', '')

    try:
        inv_date_str = datetime.strptime(date_str, '%d-%m-%Y').strftime('%d/%m/%Y')
    except:
        inv_date_str = date_str

    total_row = fill_invoice(wb, invoice_no, inv_date_str, dollar_rate, items)

    fill_packing_list(wb, invoice_no, inv_date_str, dollar_rate,
                      buyer_name, buyer_addr, items, customers)

    # PBE
    ws_pbe = wb['PBE']
    ws_pbe['G23'] = '=_xlfn.TEXTJOIN(" ",TRUE, INVOICE!I2,INVOICE!I3)'
    ws_pbe['M23'] = dollar_rate
    ws_pbe['N23'] = f'=M23*INVOICE!K{total_row}'
    ws_pbe['B29'] = '=G23'

    out = io.BytesIO()
    wb.save(out)
    out.seek(0)
    return out




# ─────────────────────────────────────────────
#  LP EXCEL GENERATOR
# ─────────────────────────────────────────────
# Country name -> code lookup (from Information sheet)
COUNTRY_CODES = {
    'AFGHANISTAN': 'AF',
    'ALBANIA': 'AL',
    'ALGERIA': 'DZ',
    'ANDORRA': 'AD',
    'ANGOLA': 'AO',
    'ANGUILLA': 'AI',
    'ANTIGUA/BARBUDA': 'AG',
    'ARGENTINA': 'AR',
    'ARMENIA': 'AM',
    'ARUBA': 'AW',
    'ASCENSION': 'SH1',
    'AUSTRALIA': 'AU',
    'AUSTRIA': 'AT',
    'AZERBAIJAN': 'AZ',
    'BAHAMAS': 'BS',
    'BAHRAIN': 'BH',
    'BANGLADESH': 'BD',
    'BARBADOS': 'BB',
    'BELARUS': 'BY',
    'BELGIUM': 'BE',
    'BELIZE': 'BZ',
    'BENIN': 'BJ',
    'BERMUDA': 'BM',
    'BHUTAN': 'BT',
    'BOLIVIA': 'BO',
    'BONAIRE, SINT EUSTATIUS AND SABA': 'BQ',
    'BOSNIA AND HERZEGOVINA': 'BA',
    'BOTSWANA': 'BW',
    'BRAZIL': 'BR',
    'BRITISH VIRGIN ISLANDS': 'VG',
    'BRUNEI DARUSSALAM': 'BN',
    'BULGARIA': 'BG',
    'BURKINA FASO': 'BF',
    'BURUNDI': 'BI',
    'CAMBODIA': 'KH',
    'CAMEROON': 'CM',
    'CANADA': 'CA',
    'CAPE VERDE': 'CV',
    'CAYMAN ISLANDS': 'KY',
    'CENTRAL AFRICAN REPUBLIC': 'CF',
    'CHAD': 'TD',
    'CHILE': 'CL',
    'CHILE EASTER ISLAND': 'CL1',
    'CHINA': 'CN',
    'COLOMBIA': 'CO',
    'COMOROS': 'KM',
    'COOK ISLANDS': 'CK',
    'COSTA RICA': 'CR',
    "COTE D'IVOIRE": 'CI',
    'CROATIA': 'HR',
    'CUBA': 'CU',
    'CURACAO': 'CW',
    'CYPRUS': 'CY',
    'CZECH REPUBLIC': 'CZ',
    'DEMOCRATIC REPUBLIC OF CONGO': 'CD',
    'DENMARK': 'DK',
    'DJIBOUTI': 'DJ',
    'DOMINICA': 'DM',
    'DOMINICAN REPUBLIC': 'DO',
    'EAST TIMOR': 'TP',
    'ECUADOR': 'EC',
    'EGYPT': 'EG',
    'EL SALVADOR': 'SV',
    'EQUATORIAL GUINEA': 'GQ',
    'ERITREA': 'ER',
    'ESTONIA': 'EE',
    'ETHIOPIA': 'ET',
    'FALKLAND ISLANDS': 'FK',
    'FAROE ISLANDS': 'FO',
    'FIJI': 'FJ',
    'FINLAND': 'FI',
    'FRANCE': 'FR',
    'FRENCH GUIANA': 'GF',
    'FRENCH POLYNESIA': 'PF',
    'FRENCH S.TERRITORY': 'TF',
    'GABON': 'GA',
    'GAMBIA': 'GM',
    'GEORGIA': 'GE',
    'GERMANY': 'DE',
    'GHANA': 'GH',
    'GIBRALTAR': 'GI',
    'GREECE': 'GR',
    'GREENLAND': 'GL',
    'GRENADA': 'GD',
    'GUADELOUPE': 'GP',
    'GUAM': 'GU',
    'GUATEMALA': 'GT',
    'GUERNSEY': 'GG',
    'GUINEA': 'GN',
    'GUINEA BISSAU': 'GW',
    'GUYANA': 'GY',
    'HAITI': 'HT',
    'HONDURAS': 'HN',
    'HONG KONG': 'HK',
    'HUNGARY': 'HU',
    'ICELAND': 'IS',
    'INDONESIA': 'ID',
    'IRAN': 'IR',
    'IRAQ': 'IQ',
    'IRELAND': 'IE',
    'ISLE OF MAN': 'IM',
    'ISRAEL': 'IL',
    'ITALY': 'IT',
    'JAMAICA': 'JM',
    'JAPAN': 'JP',
    'JERSEY': 'JE',
    'JORDAN': 'JO',
    'KAZAKHSTAN': 'KZ',
    'KENYA': 'KE',
    'KIRIBATI': 'KI',
    'KOSOVA': 'KV',
    'KUWAIT': 'KW',
    'KYRGYZSTAN': 'KG',
    'LAOS': 'LA',
    'LATVIA': 'LV',
    'LEBANON': 'LB',
    'LESOTHO': 'LS',
    'LIBERIA': 'LR',
    'LIBYA': 'LY',
    'LIEISCHENSTEIN': 'LI',
    'LITHUANIA': 'LT',
    'LUXEMBOURG': 'LU',
    'MACAU': 'MO',
    'MACEDONIA': 'MK',
    'MADAGASCAR': 'MG',
    'MALAWI': 'MW',
    'MALAYSIA': 'MY',
    'MALDIVES': 'MV',
    'MALI': 'ML',
    'MALTA': 'MT',
    'MARSHALL ISLANDS': 'MH',
    'MARTINIQUE': 'MQ',
    'MAURITANIA': 'MR',
    'MAURITIUS': 'MU',
    'MAYOTTE': 'YT',
    'MEXICO': 'MX',
    'MICRONESIA': 'FM',
    'MOLDOVA': 'MD',
    'MONACO': 'MC',
    'MONGOLIA': 'MN',
    'MONTENEGRO': 'ME',
    'MONTSERRAT': 'MS',
    'MOROCCO': 'MA',
    'MOZAMBIQUE': 'MZ',
    'MYANMAR': 'MM',
    'NAMIBIA': 'NA',
    'NAURU': 'NR',
    'NEPAL': 'NP',
    'NETHERLANDS': 'NL',
    'NEW CALEDONIA': 'NC',
    'NEW ZEALAND': 'NZ',
    'NICARAGUA': 'NI',
    'NIGER': 'NE',
    'NIGERIA': 'NG',
    'NIUE': 'NU',
    'NORTH KOREA': 'KP',
    'NORWAY': 'NO',
    'OMAN': 'OM',
    'PAKISTAN': 'PK',
    'PALAU': 'PW',
    'PALESTINE': 'PS',
    'PANAMA': 'PA',
    'PAPUA NEW GUINEA': 'PG',
    'PARAGUAY': 'PY',
    'PERU': 'PE',
    'PHILIPPINES': 'PH',
    'PITCAIRN ISLANDS': 'PN',
    'POLAND': 'PL',
    'PORTUGAL': 'PT',
    'PUERTO RICO': 'PR',
    'QATAR': 'QA',
    'REPUBLIC OF CONGO': 'CG',
    'REUNION': 'RE',
    'ROMANIA': 'RO',
    'RUSSIAN FED': 'RU',
    'RUSSIAN (MOSCOW)': 'RU1',
    'RUSSIAN (ST. PETERSBERG)': 'RU2',
    'RWANDA': 'RW',
    'SAINT HELENA': 'SH',
    'SAINT PIERRE AND MIQUELON': 'PM',
    'SAIPAN': 'MP1',
    'SAMOA': 'WS',
    'SAMOA, AMERICA': 'AS',
    'SAN MARINO': 'SM',
    'SAUDI ARABIA': 'SA',
    'SENEGAL': 'SN',
    'SERBIA': 'RS',
    'SERBIA AND MONTENEGRO': 'CS',
    'SEYCHELLES': 'SC',
    'SIERRA LEONE': 'SL',
    'SINGAPORE': 'SG',
    'SINT MAARTEN': 'SX',
    'SLOVAKIA': 'SK',
    'SLOVENIA': 'SI',
    'SOLOMON ISLANDS': 'SB',
    'SOMALIA': 'SO',
    'SOMALILAND': 'SO1',
    'SOUTH AFRICA': 'ZA',
    'SOUTH KOREA': 'KR',
    'SOUTH SUDAN': 'SS',
    'SPAIN': 'ES',
    'SPAIN(CANARY ISLANDS)': 'ES1',
    'SRI LANKA': 'LK',
    'ST. BARTHELEMY': 'BL',
    'ST KITTS & NEVIS': 'KN',
    'ST. LUCIA': 'LC',
    'S.TOME,PRINCIPE': 'ST',
    'ST. VINCENT & THE GRENADINES': 'VC',
    'SUDAN': 'SD',
    'SURINAME': 'SR',
    'SVALBARD AND JAN MAYEN': 'SJ',
    'SWAZILAND': 'SZ',
    'SWEDEN': 'SE',
    'SWITZERLAND': 'CH',
    'SYRIA': 'SY',
    'TAHITI': 'PF1',
    'TAIWAN': 'TW',
    'TAJIKISTAN': 'TJ',
    'TANZANIA': 'TZ',
    'THAILAND': 'TH',
    'TIMOR-LESTE': 'TL1',
    'TOGO': 'TG',
    'TONGA': 'TO',
    'TRINIDAD AND TOBAGO': 'TT',
    'TRISTAN DA CUNHA': 'SH2',
    'TUNISIA': 'TN',
    'TURKEY': 'TR',
    'TURKMENISTAN': 'TM',
    'TURKS AND CAICOS ISLANDS': 'TC',
    'TUVALU': 'TV',
    'UGANDA': 'UG',
    'UKRAINE': 'UA',
    'UNITED ARAB EMIRATES': 'AE',
    'UNITED STATES': 'US',
    'UNITED KINGDOM': 'GB',
    'UK': 'GB',
    'UNITED STATES OF AMERICA': 'US',
    'USA': 'US',
    'URUGUAY': 'UY',
    'U.S. VIRGIN ISLANDS': 'VI',
    'UZBEKISTAN': 'UZ',
    'VANUATU': 'VU',
    'VATICAN CITY': 'VA',
    'VENEZUELA': 'VE',
    'VIETNAM': 'VN',
    'WALLIS AND FUTUNA': 'WF',
    'WESTERN SAHARA': 'EH',
    'YEMEN': 'YE',
    'ZAMBIA': 'ZM',
    'ZIMBABWE': 'ZW',
}

LP_TEMPLATE_PATH = os.path.join(os.path.dirname(__file__), 'lp_template.xlsx')

def generate_lp_excel(invoice_no, customers):
    """
    Generate the LP (Label Print) Excel.
    One row per customer.
    Dynamic cols: A(serial), B(country code via formula), C(country name),
                  G(weight), H(declared value), X(receiver name),
                  Z(addr1), AA(addr2 blank), AC(city), AD(state), AE(zip),
                  AJ(tracking)
    All other columns: hardcoded from template.
    """
    wb = load_workbook(LP_TEMPLATE_PATH)
    ws = wb['ArticleDetails']

    # Clear ALL existing data rows (rows 2 onwards) so old template data never bleeds through
    for row in ws.iter_rows(min_row=2, max_row=ws.max_row):
        for cell in row:
            cell.value = None

    # Hardcoded sender / static values (same every row)
    STATIC = {
        5:  999,                                  # E  MAIL NATURE TYPE
        6:  'AMS',                                # F  MAIL TRANSPORT TYPE
        9:  'P',                                  # I  NON DELIVERY INSTRUCTIONS
        10: 'NARIX ',                             # J  SENDER NAME
        11: 'NARIX PAHRMACEUTICALS PVT LTD',      # K  SENDER COMPANY
        12: '404, 4TH FLOOR, BAYVUE EAST',        # L  SENDER ADD LINE 1
        13: 'DR. MB VELKAR STREET, KALBADEVI',    # M  SENDER ADD LINE 2
        15: 'MUMBAI',                             # O  SENDER CITY
        16: 'Maharashtra',                        # P  SENDER STATE/UT
        17: 'INDIA',                              # Q  SENDER COUNTRY NAME
        18: 'IN',                                 # R  SENDER COUNTRY CODE
        19: 400002,                               # S  SENDER PINCODE
        35: False,                                # AI ALT ADD FLAG
        37: 1234567890,                           # AK SENDER MOBILE
        38: 1234567890,                           # AL RECEIVER MOBILE
        43: False,                                # AQ POD FLAG
        45: 'PBEII',                              # AS PBE TYPE
        47: 30049099,                             # AU HS CODE
        48: 'MEDICINE',                           # AV HS DESCRIPTION
    }

    for row_idx, cust in enumerate(customers):
        r = row_idx + 2   # data starts at row 2

        # A — Serial number
        ws.cell(r, 1).value = row_idx + 1

        # B — Country code via VLOOKUP formula (auto-resolves from C)
        ws.cell(r, 2).value = f'=IF(C{r}="","", VLOOKUP(C{r}, Information!$N$2:$O$251, 2, FALSE))'

        # C — Destination country name
        country_name = str(cust.get('country', '')).strip().upper()
        ws.cell(r, 3).value = country_name

        # G — Physical weight
        ws.cell(r, 7).value = float(cust.get('weight', 0))

        # H — Declared value
        ws.cell(r, 8).value = float(cust.get('declared_value', 0))

        # X — Receiver name
        ws.cell(r, 24).value = str(cust.get('name', '')).strip().upper()

        # Z — Receiver address line 1
        ws.cell(r, 26).value = str(cust.get('addr1', '')).strip().upper()

        # AA — Receiver address line 2 (blank — we only have addr1)
        ws.cell(r, 27).value = str(cust.get('addr2', '')).strip().upper()

        # AC — Receiver city
        ws.cell(r, 29).value = str(cust.get('city', '')).strip().upper()

        # AD — Receiver state
        ws.cell(r, 30).value = str(cust.get('state', '')).strip().upper()

        # AE — Receiver zip/postcode
        ws.cell(r, 31).value = str(cust.get('zip', '')).strip()

        # AJ — Article number / tracking (optional)
        tracking = str(cust.get('tracking', '')).strip()
        ws.cell(r, 36).value = tracking if tracking else None

        # All static/hardcoded columns
        for col, val in STATIC.items():
            ws.cell(r, col).value = val

    out = io.BytesIO()
    wb.save(out)
    out.seek(0)
    return out
# ─────────────────────────────────────────────
#  ROUTES
# ─────────────────────────────────────────────
@app.route('/health')
def health():
    return 'ok', 200

@app.route('/product-db')
def product_db_route():
    return jsonify(PRODUCT_DB)

@app.route('/')
def index():
    return send_from_directory('static', 'index.html')

@app.route('/parse-pdf', methods=['POST'])
def parse_pdf_route():
    if 'pdf' not in request.files:
        return jsonify({'error': 'No PDF uploaded'}), 400
    return jsonify(parse_pdf(request.files['pdf'].read()))

@app.route('/generate-excel', methods=['POST'])
def generate_excel_route():
    data        = request.json
    dollar_rate = data.get('dollar_rate')
    items       = data.get('items', [])
    customers   = data.get('customers', [])

    errors = []
    if not dollar_rate or float(dollar_rate) <= 0:
        errors.append('Dollar rate is required.')
    elif float(dollar_rate) > 200:
        errors.append(f'Dollar rate {dollar_rate} looks wrong — must be between 1 and 200.')
    for idx, item in enumerate(items):
        if not item.get('generic', '').strip():
            errors.append(f'Row {idx+1} ({item.get("product","")}): Generic name is required.')
        if not item.get('manufacturer', '').strip():
            errors.append(f'Row {idx+1} ({item.get("product","")}): Manufacturer is required.')
    if not customers:
        errors.append('At least one customer is required in the Packing List.')
    for ci, cust in enumerate(customers):
        n = ci + 1
        if not cust.get('name', '').strip():
            errors.append(f'Customer {n}: Receiver Name is required.')
        if not cust.get('addr1', '').strip():
            errors.append(f'Customer {n}: Address Line 1 is required.')
        if not cust.get('city', '').strip():
            errors.append(f'Customer {n}: City is required.')
        if not cust.get('state', '').strip():
            errors.append(f'Customer {n}: State is required.')
        if not str(cust.get('zip', '')).strip():
            errors.append(f'Customer {n}: ZIP / Postcode is required.')
        if not cust.get('country', '').strip():
            errors.append(f'Customer {n}: Country is required.')
        if not cust.get('lines'):
            errors.append(f'Customer {n}: At least one product line is required.')
    if errors:
        return jsonify({'errors': errors}), 422

    try:
        excel  = generate_excel(data)
    except Exception as e:
        import traceback
        return jsonify({'errors': [f'Server error: {str(e)}']}), 500

    inv_no = data.get('invoice_no', 'invoice').replace('/', '-')
    return send_file(excel,
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        as_attachment=True,
        download_name=f'{inv_no}.xlsx')


@app.route('/generate-lp', methods=['POST'])
def generate_lp_route():
    data      = request.json
    customers = data.get('customers', [])
    invoice_no = data.get('invoice_no', 'invoice')

    errors = []
    if not customers:
        errors.append('At least one customer is required.')
    for ci, cust in enumerate(customers):
        n = ci + 1
        if not cust.get('name', '').strip():
            errors.append(f'Customer {n}: Receiver Name is required.')
        if not cust.get('addr1', '').strip():
            errors.append(f'Customer {n}: Address Line 1 is required.')
        if not cust.get('city', '').strip():
            errors.append(f'Customer {n}: City is required.')
        if not cust.get('state', '').strip():
            errors.append(f'Customer {n}: State is required.')
        if not str(cust.get('zip', '')).strip():
            errors.append(f'Customer {n}: ZIP / Postcode is required.')
        if not cust.get('country', '').strip():
            errors.append(f'Customer {n}: Country is required.')
        # weight=0 is accepted (some shipments have no declared weight)
        # No weight validation — field is optional
        # declared_value=0 is accepted
        if cust.get('declared_value') is None:
            pass  # 0 is fine
    if errors:
        return jsonify({'errors': errors}), 422

    try:
        lp_excel  = generate_lp_excel(invoice_no, customers)
    except Exception as e:
        return jsonify({'errors': [f'Server error: {str(e)}']}), 500

    inv_no    = invoice_no.replace('/', '-')
    return send_file(lp_excel,
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        as_attachment=True,
        download_name=f'{inv_no}_LP.xlsx')

if __name__ == '__main__':
    import logging
    logging.getLogger('werkzeug').setLevel(logging.WARNING)
    print("  Server ready -> http://localhost:5000")
    app.run(debug=False, port=5000, use_reloader=False)
