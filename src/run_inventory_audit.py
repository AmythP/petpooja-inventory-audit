"""
Automated Inventory Reconciliation & Stock Variance Audit System
Integrates POS/ERP consumption reports with physical counts and stock ledgers.
"""
from __future__ import annotations
import os
import re
import sys
import shutil
import datetime as dt
from html.parser import HTMLParser
from collections import defaultdict
import openpyxl
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.formatting.rule import FormulaRule

# Directory resolution: uses working dir or relative project root
BASE = os.environ.get("INVENTORY_AUDIT_DIR", os.path.dirname(os.path.abspath(__file__)))
OUTPUT = os.path.join(BASE, "production+closing=yersterday.xlsx")
STOCK_FILE = os.path.join(BASE, "sweetness available stock today.xlsx")

DEFAULT_ITEMS = [
    'Vanleer Dark Chocolate', 'Vanleer Milk Chocolate', 'Vanleer White Chocolate',
    'Dark Chocochips', 'Milk Chocochips', 'White Chocochips', 'Walnut 250gram', 'Wet Yeast',
    'Mithai Mate 400 Gm', 'Pista 500gm', 'Almond', 'Gems 500grm', 'Badam Flake',
    'Kitkat 2 24(42*11.9)', 'Butter 500gm I/p Amul', 'Delicious Butter',
    'Dark Chocolate Ganache', 'Milk Chocolate Ganache', 'White Chocolate Ganache'
]

NAVY = '1F4E79'
BLUE = 'DDEBF7'
YELLOW = 'FFF2CC'
GREEN = 'E2F0D9'
RED = 'FCE4D6'

header_fill = PatternFill('solid', fgColor=NAVY)
input_fill = PatternFill('solid', fgColor=YELLOW)
section_fill = PatternFill('solid', fgColor=BLUE)
bad_fill = PatternFill('solid', fgColor=RED)
white_bold = Font(color='FFFFFF', bold=True)
bold = Font(bold=True)
thin = Side(style='thin', color='B7B7B7')
border = Border(left=thin, right=thin, top=thin, bottom=thin)
center = Alignment(horizontal='center', vertical='center', wrap_text=True)


class PetpoojaHTMLParser(HTMLParser):
    def __init__(self):
        super().__init__()
        self.rows = []
        self.row = []
        self.cell = []
        self.inside = False

    def handle_starttag(self, tag, attrs):
        if tag == 'tr':
            self.row = []
        elif tag in ('td', 'th'):
            self.inside = True
            self.cell = []

    def handle_data(self, data):
        if self.inside:
            self.cell.append(data)

    def handle_endtag(self, tag):
        if tag in ('td', 'th'):
            self.inside = False
            self.row.append(''.join(self.cell).strip())
        elif tag == 'tr' and self.row:
            self.rows.append(self.row)


def clean_number(v):
    if v is None:
        return 0.0
    if isinstance(v, (int, float)):
        return float(v)
    m = re.search(r'-?[\d,]+(?:\.\d+)?', str(v))
    return float(m.group(0).replace(',', '')) if m else 0.0


def normalize_name(x):
    return re.sub(r'[^a-z0-9]', '', str(x).lower())


def get_consumption_file():
    candidates = []
    for name in os.listdir(BASE):
        low = name.lower()
        if 'consumption' in low and low.endswith('.xls') and not low.endswith('.xlsx'):
            path = os.path.join(BASE, name)
            candidates.append(path)
    if not candidates:
        raise FileNotFoundError('No consumption .xls file found in audit directory.')
    return max(candidates, key=os.path.getmtime)


def parse_consumption(path):
    with open(path, 'r', encoding='utf-8-sig', errors='ignore') as f:
        raw = f.read()

    # Handle HTML frameset exports
    if 'sheet001.htm' in raw.lower():
        folder = os.path.splitext(path)[0] + '_files'
        sheet = os.path.join(folder, 'sheet001.htm')
        if os.path.exists(sheet):
            with open(sheet, 'r', encoding='utf-8-sig', errors='ignore') as f:
                raw = f.read()

    parser = PetpoojaHTMLParser()
    parser.feed(raw)

    header = None
    for ix, row in enumerate(parser.rows):
        cells = [str(x).strip().lower() for x in row]
        if any('raw material' in x for x in cells):
            header = ix
            break

    if header is None:
        raise ValueError('Consumption report header "Raw material" not found.')

    totals = defaultdict(float)
    raw_rows = []
    report_date = None

    for row in parser.rows:
        text = ' '.join(row)
        m = re.search(r'\[(\d{1,2}\s+[A-Za-z]{3}\s+\d{2,4})\]', text)
        if m:
            for fmt in ('%d %b %Y', '%d %b %y'):
                try:
                    report_date = dt.datetime.strptime(m.group(1), fmt).date()
                    break
                except ValueError:
                    pass

    for row in parser.rows[header + 1:]:
        if len(row) < 10 or 'total order' in ' '.join(row).lower() or 'grand total' in ' '.join(row).lower():
            continue
        material = str(row[8]).strip()
        qty = clean_number(row[9])
        unit = str(row[10]).strip() if len(row) > 10 else ''

        if not unit:
            unit = re.sub(r'^\s*-?[\d,.]+', '', str(row[9])).replace("'", '').strip()

        if material and qty:
            totals[material] += qty
            raw_rows.append((material, qty, unit))

    if not totals:
        raise ValueError('No consumption rows were parsed from report.')

    return dict(totals), raw_rows, report_date or dt.date.today()


def parse_stock(path):
    wb = openpyxl.load_workbook(path, data_only=True)
    ws = wb.active
    data = []
    lookup = {}

    for r in range(2, ws.max_row + 1):
        name = ws.cell(r, 2).value
        if not name:
            continue
        row = {
            'category': str(ws.cell(r, 1).value or ''),
            'name': str(name).strip(),
            'available': clean_number(ws.cell(r, 3).value),
            'update': clean_number(ws.cell(r, 4).value),
            'unit': str(ws.cell(r, 5).value or ''),
            'price_old': ws.cell(r, 6).value,
            'price_new': ws.cell(r, 7).value,
            'comments': ws.cell(r, 8).value,
            'barcode': ws.cell(r, 9).value
        }
        data.append(row)
        lookup[normalize_name(row['name'])] = row

    return data, lookup


def map_item(item, stock, cons):
    ni = normalize_name(item)
    for name in list(stock) + list(cons):
        if normalize_name(name) == ni:
            return name
    for name in list(stock) + list(cons):
        nn = normalize_name(name)
        if ni in nn or nn in ni:
            return name
    aliases = {'mithaimate400gm': 'Mithai Mate 400'}
    target = aliases.get(ni)
    if target:
        for name in list(stock) + list(cons):
            if normalize_name(target) in normalize_name(name):
                return name
    return item


def prior_state(report_date):
    opening = {}
    same_purchase = {}
    same_adjust = {}
    same_actual = {}
    same_reason = {}
    history = []
    same_day = False
    items = []

    if not os.path.exists(OUTPUT):
        return opening, same_purchase, same_adjust, same_actual, same_reason, history, same_day, items

    try:
        values_wb = openpyxl.load_workbook(OUTPUT, data_only=True)
        has_meta = 'META' in values_wb.sheetnames
        if has_meta and values_wb['META']['B1'].value == report_date.isoformat():
            same_day = True
        if not has_meta:
            same_day = True

        if 'Sheet1' in values_wb.sheetnames:
            s = values_wb['Sheet1']
            items = [s.cell(r, 1).value for r in range(2, s.max_row + 1) if s.cell(r, 1).value]
            for r in range(2, s.max_row + 1):
                if s.cell(r, 1).value:
                    same_actual[str(s.cell(r, 1).value)] = clean_number(s.cell(r, 3).value)
                    if s.max_column >= 12:
                        same_reason[str(s.cell(r, 1).value)] = s.cell(r, 12).value or ''

        if 'DAILY CHECK' in values_wb.sheetnames:
            s = values_wb['DAILY CHECK']
            for r in range(2, s.max_row + 1):
                item = s.cell(r, 1).value
                if not item:
                    continue
                item = str(item)
                if same_day:
                    opening[item] = clean_number(s.cell(r, 2).value)
                    same_purchase[item] = clean_number(s.cell(r, 3).value)
                    same_adjust[item] = clean_number(s.cell(r, 8).value)
                else:
                    opening[item] = clean_number(s.cell(r, 5).value) or same_actual.get(item, 0)

        if 'HISTORY' in values_wb.sheetnames:
            s = values_wb['HISTORY']
            for row in s.iter_rows(min_row=2, values_only=True):
                if any(v is not None for v in row):
                    history.append(list(row))

        if not same_day and 'DAILY CHECK' in values_wb.sheetnames and 'META' in values_wb.sheetnames:
            old_date = values_wb['META']['B1'].value or ''
            s = values_wb['DAILY CHECK']
            for r in range(2, s.max_row + 1):
                if s.cell(r, 1).value:
                    history.append([old_date] + [s.cell(r, c).value for c in range(1, 12)])
    except Exception as e:
        print(f'WARNING: Prior state read failure: {e}')

    return opening, same_purchase, same_adjust, same_actual, same_reason, history, same_day, items


def style_table(ws, start, end, cols):
    for r in range(start, end + 1):
        for c in range(1, cols + 1):
            cell = ws.cell(r, c)
            cell.border = border
            if r == start:
                cell.fill = header_fill
                cell.font = white_bold
                cell.alignment = center
            else:
                cell.alignment = Alignment(vertical='center')


def set_widths(ws, widths):
    for col, w in widths.items():
        ws.column_dimensions[col].width = w


def run_audit():
    cons_file = get_consumption_file()
    consumption, raw_rows, report_date = parse_consumption(cons_file)
    if not os.path.exists(STOCK_FILE):
        raise FileNotFoundError(f'Missing stock file: {STOCK_FILE}')

    stock_rows, stock = parse_stock(STOCK_FILE)
    opening, preserve_purchase, preserve_adjust, preserve_actual, preserve_reason, history, same_day, old_items = prior_state(report_date)
    items = old_items or DEFAULT_ITEMS
    items = list(dict.fromkeys(str(x) for x in items))
    mapping = {item: map_item(item, stock, consumption) for item in items}

    if os.path.exists(OUTPUT):
        backup = os.path.join(BASE, f'backup_inventory_{dt.datetime.now():%Y%m%d_%H%M%S}.xlsx')
        shutil.copy2(OUTPUT, backup)

    wb = Workbook()

    # Sheet1: Core Audit Ledger
    ws = wb.active
    ws.title = 'Sheet1'
    headers = [
        'ITEM', "TODAY'S PRODUCTION", "TODAY'S CLOSING\n(Physical Count)",
        'EXPECTED OPENING', 'ACTUAL OPENING', 'DIFF. OPENING',
        'EXPECTED CLOSING', 'PETPUJA CLOSING', 'INVENTORY MISMATCH',
        'PURCHASE / TRANSFER IN', 'ADJUSTMENT', 'REASON FOR MISMATCH'
    ]
    ws.append(headers)
    ws.row_dimensions[1].height = 34

    for r, item in enumerate(items, 2):
        mapped = mapping[item]
        srow = stock.get(normalize_name(mapped), {})
        prod = consumption.get(mapped, consumption.get(item, 0))

        ws.cell(r, 1, item)
        ws.cell(r, 2, round(prod, 3))
        ws.cell(r, 3, preserve_actual.get(item, 0) if same_day else 0)
        ws.cell(r, 4, f"='DAILY CHECK'!B{r}")
        ws.cell(r, 5, opening.get(item, 0))
        ws.cell(r, 6, f'=D{r}-E{r}')
        ws.cell(r, 7, f'=D{r}+J{r}-B{r}+K{r}')
        ws.cell(r, 8, srow.get('available', 0))
        ws.cell(r, 9, f'=C{r}-G{r}')
        ws.cell(r, 10, f"='DAILY CHECK'!C{r}")
        ws.cell(r, 11, f"='DAILY CHECK'!H{r}")
        ws.cell(r, 12, f'=IF(ABS(I{r})<=5,"Matched",IF(I{r}<0,"Shortage of "&ROUND(ABS(I{r}),2)&" (Wastage/Chori)","Excess of "&ROUND(I{r},2)&" (Missed Inward/Recipe)"))')

        for c in range(2, 12):
            ws.cell(r, c).number_format = '#,##0.00'
        ws.cell(r, 1).font = bold
        ws.cell(r, 3).fill = input_fill

    style_table(ws, 1, len(items) + 1, len(headers))
    set_widths(ws, {
        'A': 29, 'B': 17, 'C': 18, 'D': 18, 'E': 16, 'F': 14,
        'G': 18, 'H': 17, 'I': 19, 'J': 20, 'K': 15, 'L': 40
    })
    ws.freeze_panes = 'B2'
    ws.auto_filter.ref = f'A1:L{len(items) + 1}'
    ws.conditional_formatting.add(f'I2:I{len(items) + 1}', FormulaRule(formula=['ABS(I2)>5'], fill=bad_fill))

    # DAILY CHECK Tab
    dc = wb.create_sheet('DAILY CHECK')
    prev = (report_date - dt.timedelta(days=1)).strftime('%d-%b')
    now = report_date.strftime('%d-%b')
    dc_headers = [
        'ITEM', f'Opening Closing {prev}', 'Purchase / Transfer In',
        f'Petpuja Closing {now}', f'Actual Closing {now}', 'Diff. Petpuja vs Actual',
        'Remark', 'Warmer / Other Adjustment', "Today's Production",
        'Expected Net Closing', 'Shortage / Excess'
    ]
    dc.append(dc_headers)
    dc.row_dimensions[1].height = 34

    for r, item in enumerate(items, 2):
        dc.cell(r, 1, item)
        dc.cell(r, 2, opening.get(item, 0))
        dc.cell(r, 3, preserve_purchase.get(item, 0) if same_day else 0)
        dc.cell(r, 4, f'=Sheet1!H{r}')
        dc.cell(r, 5, f'=Sheet1!C{r}')
        dc.cell(r, 6, f'=D{r}-E{r}')
        dc.cell(r, 7, '')
        dc.cell(r, 8, preserve_adjust.get(item, 0) if same_day else 0)
        dc.cell(r, 9, f'=Sheet1!B{r}')
        dc.cell(r, 10, f'=B{r}+C{r}-I{r}-H{r}')
        dc.cell(r, 11, f'=J{r}-E{r}')

        for c in [2, 3, 4, 5, 6, 8, 9, 10, 11]:
            dc.cell(r, c).number_format = '#,##0.00'
        dc.cell(r, 1).font = bold
        dc.cell(r, 3).fill = input_fill
        dc.cell(r, 8).fill = input_fill

    style_table(dc, 1, len(items) + 1, 11)
    set_widths(dc, {
        'A': 29, 'B': 18, 'C': 20, 'D': 18, 'E': 18, 'F': 20,
        'G': 28, 'H': 23, 'I': 17, 'J': 20, 'K': 18
    })
    dc.freeze_panes = 'B2'
    dc.auto_filter.ref = f'A1:K{len(items) + 1}'
    dc.conditional_formatting.add(f'K2:K{len(items) + 1}', FormulaRule(formula=['ABS(K2)>5'], fill=bad_fill))

    # Raw Production tab
    pr = wb.create_sheet('refresh today production')
    pr.append(['Raw Material', 'Qty', 'Unit', '', '', 'Row Labels', 'Sum of Qty'])
    style_table(pr, 1, 1, 7)
    for rr, (mat, qty, unit) in enumerate(raw_rows, 2):
        pr.cell(rr, 1, mat)
        pr.cell(rr, 2, qty)
        pr.cell(rr, 3, unit)
        pr.cell(rr, 2).number_format = '#,##0.000'
    for rr, (mat, qty) in enumerate(sorted(consumption.items()), 2):
        pr.cell(rr, 6, mat)
        pr.cell(rr, 7, qty)
        pr.cell(rr, 7).number_format = '#,##0.000'
    set_widths(pr, {'A': 30, 'B': 13, 'C': 12, 'F': 30, 'G': 16})
    pr.freeze_panes = 'A2'

    # Available Stock tab
    st = wb.create_sheet('sweetness available stock')
    st.append(['Category', 'Raw Material', 'Available Stock', 'Update Stock', 'Unit', 'Existing Avg Price', 'Current Avg Price', 'Comments', 'Barcode'])
    for d in stock_rows:
        st.append([d['category'], d['name'], d['available'], d['update'], d['unit'], d['price_old'], d['price_new'], d['comments'], d['barcode']])
    style_table(st, 1, len(stock_rows) + 1, 9)
    set_widths(st, {'A': 16, 'B': 31, 'C': 16, 'D': 16, 'E': 10, 'F': 18, 'G': 18, 'H': 24, 'I': 20})
    st.freeze_panes = 'B2'
    st.auto_filter.ref = f'A1:I{len(stock_rows) + 1}'

    # Executive Print Sheet
    fp = wb.create_sheet('FOR PRINT')
    fp.append(['KITCHEN DAILY INVENTORY AUDIT', report_date.strftime('%d-%b-%Y'), '', '', ''])
    fp.merge_cells('A1:E1')
    fp['A1'].font = Font(bold=True, size=14, color='FFFFFF')
    fp['A1'].fill = header_fill
    fp['A1'].alignment = center
    fp.append(['ITEM', 'OPENING', 'PURCHASE IN', 'PRODUCTION', 'PHYSICAL CLOSING'])
    style_table(fp, 2, 2, 5)
    for r, item in enumerate(items, 3):
        src = r - 1
        fp.cell(r, 1, item)
        fp.cell(r, 2, f"='DAILY CHECK'!B{src}")
        fp.cell(r, 3, f"='DAILY CHECK'!C{src}")
        fp.cell(r, 4, f'=Sheet1!B{src}')
        fp.cell(r, 5, f'=Sheet1!C{src}')
        for c in range(2, 6):
            fp.cell(r, c).number_format = '#,##0.00'
    style_table(fp, 2, len(items) + 2, 5)
    set_widths(fp, {'A': 32, 'B': 16, 'C': 16, 'D': 16, 'E': 18})
    fp.freeze_panes = 'A3'

    # Historical Tracking
    hs = wb.create_sheet('HISTORY')
    hs.append(['Report Date'] + dc_headers)
    for row in history:
        hs.append(row)
    style_table(hs, 1, max(1, hs.max_row), len(dc_headers) + 1)
    set_widths(hs, {'A': 14, 'B': 29, 'C': 16, 'D': 18, 'E': 18, 'F': 18, 'G': 18, 'H': 28, 'I': 23, 'J': 17, 'K': 20, 'L': 18})
    hs.freeze_panes = 'A2'

    # Metadata Sheet
    meta = wb.create_sheet('META')
    meta['A1'] = 'Report Date'
    meta['B1'] = report_date.isoformat()
    meta['A2'] = 'Consumption Source'
    meta['B2'] = os.path.basename(cons_file)
    meta['A3'] = 'Run Mode'
    meta['B3'] = 'Same-day refresh' if same_day else 'New day rollover'
    meta.sheet_state = 'hidden'

    wb.calculation.fullCalcOnLoad = True
    wb.calculation.forceFullCalc = True
    wb.calculation.calcMode = 'auto'
    wb.save(OUTPUT)
    print(f'Successfully built audit workbook: {OUTPUT}')


if __name__ == '__main__':
    try:
        run_audit()
    except Exception as e:
        print(f'Audit execution failed: {e}', file=sys.stderr)
        sys.exit(1)
