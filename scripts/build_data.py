#!/usr/bin/env python3
"""
Multi-Agent Verified Data Pipeline for Stock list.xlsx & NOV. SHIPMENT -OCT22.xlsx
Strictly extracts Warehouse 01 & Warehouse 60 inventory, pairs Left/Right (LHF/RHF) items,
and reconciles against Monthly (NOV-full), Weekly (OCT-WK 42), and Master Opening (Data) plans.
"""
import zipfile
import xml.etree.ElementTree as ET
import json
import os
import re
import shutil
from collections import defaultdict
from datetime import datetime, timedelta, timezone

ROOT_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

def clean_temp():
    audit_dir = os.path.join(ROOT_DIR, '.audit')
    if os.path.exists(audit_dir):
        shutil.rmtree(audit_dir, ignore_errors=True)

def excel_date(val):
    if val is None:
        return ''
    s = str(val).strip()
    if not s or s == 'None' or s.startswith('#'):
        return ''
    try:
        n = float(s)
        if 40000 <= n <= 55000:
            dt = datetime(1899, 12, 30) + timedelta(days=int(n))
            return dt.strftime('%Y-%m-%d')
    except Exception:
        pass
    return s

def safe_float(val, default=0.0):
    if val is None:
        return default
    try:
        return float(str(val).strip())
    except Exception:
        return default

def parse_xlsx(path):
    ns = {
        'main': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main',
        'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
    }
    with zipfile.ZipFile(path) as z:
        shared_strings = []
        if 'xl/sharedStrings.xml' in z.namelist():
            root = ET.fromstring(z.read('xl/sharedStrings.xml'))
            for si in root.findall('main:si', ns):
                texts = [t.text or '' for t in si.findall('.//main:t', ns)]
                shared_strings.append(''.join(texts))

        wb_root = ET.fromstring(z.read('xl/workbook.xml'))
        rels_root = ET.fromstring(z.read('xl/_rels/workbook.xml.rels'))
        rel_map = {rel.attrib['Id']: rel.attrib['Target'] for rel in rels_root}

        sheets = {}
        for sh in wb_root.findall('.//main:sheet', ns):
            name = sh.attrib['name']
            rid = sh.attrib['{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id']
            target = rel_map[rid]
            if not target.startswith('xl/'):
                target = 'xl/' + target.lstrip('/')
            if target not in z.namelist():
                continue
            sh_root = ET.fromstring(z.read(target))
            rows_data = []
            for row in sh_root.findall('.//main:sheetData/main:row', ns):
                r_idx = int(row.attrib.get('r', 0))
                row_dict = {'_row': r_idx}
                for c in row.findall('main:c', ns):
                    ref = c.attrib.get('r', '')
                    col_letters = re.sub(r'[0-9]', '', ref)
                    t = c.attrib.get('t', '')
                    v_el = c.find('main:v', ns)
                    val = v_el.text if v_el is not None else None
                    if t == 's' and val is not None:
                        val = shared_strings[int(val)]
                    elif t == 'inlineStr':
                        is_el = c.find('.//main:t', ns)
                        val = is_el.text if is_el is not None else None
                    row_dict[col_letters] = val
                rows_data.append(row_dict)
            sheets[name] = rows_data
        return sheets

def detect_side(code, description=''):
    desc_up = (description or '').upper()
    # Finished goods 8... (10 digits) -> 6th char '1' is Left (LHF), '2' is Right (RHF)
    if len(code) == 10 and code.startswith('8'):
        if code[5] == '1':
            return 'L'
        elif code[5] == '2':
            return 'R'
    if any(k in desc_up for k in ['LHF', 'LEFT', 'GLD SIDE L', '333L', '334L', '323L', '633L', '634L']):
        return 'L'
    if any(k in desc_up for k in ['RHF', 'RIGHT', 'GLD SIDE R', '333R', '334R', '323R', '633R', '634R']):
        return 'R'
    # Explicit known pairs
    known_left = {'1220190004', '1421710201', '1427790001', '1455350001'}
    known_right = {'1220200004', '1421720201', '1427800001', '1455360001'}
    if code in known_left:
        return 'L'
    if code in known_right:
        return 'R'
    return 'N'

def build_dataset():
    clean_temp()
    stock_path = os.path.join(ROOT_DIR, 'Stock list.xlsx')
    ship_path = os.path.join(ROOT_DIR, 'NOV. SHIPMENT -OCT22.xlsx')

    stock_sheets = parse_xlsx(stock_path)
    ship_sheets = parse_xlsx(ship_path)

    # 1. Extract metadata (Feature & Description) from Shipment sheets
    item_meta = {}
    for sname in ['Data', 'NOV-full', 'OCT-WK 42']:
        if sname not in ship_sheets:
            continue
        for r in ship_sheets[sname][1:]:
            code = str(r.get('D' if sname == 'Data' else 'F') or '').strip()
            feat = str(r.get('E' if sname == 'Data' else 'H') or '').strip()
            desc = str(r.get('F' if sname == 'Data' else 'I') or '').strip()
            cust_code = '' if sname == 'Data' else str(r.get('G') or '').strip()
            if code and code != 'None':
                if code not in item_meta or not item_meta[code].get('description'):
                    item_meta[code] = {
                        'feature': feat if feat != 'None' else '',
                        'description': desc if desc != 'None' else '',
                        'customerCode': cust_code if cust_code != 'None' else ''
                    }

    # 2. Parse Stock list.xlsx (Strictly filter Warehouse '01' and '60')
    stock_rows = stock_sheets.get('Stock Balance With Batch', [])[1:]
    total_raw_batches = len(stock_rows)
    wh_summary = defaultdict(lambda: {'rows': 0, 'qty': 0.0, 'codes': set()})

    stock_by_code = defaultdict(lambda: {
        'code': '',
        'wh01Qty': 0.0,
        'wh60Qty': 0.0,
        'totalQty': 0.0,
        'batches01': 0,
        'batches60': 0,
        'bins': set(),
        'latestDate': '',
        'batchList': []
    })

    for r in stock_rows:
        wh = str(r.get('B') or '').strip()
        code = str(r.get('A') or '').strip()
        qty = safe_float(r.get('F'), 0.0)
        if not code:
            continue
        wh_summary[wh]['rows'] += 1
        wh_summary[wh]['qty'] += qty
        wh_summary[wh]['codes'].add(code)

        if wh not in ('01', '60'):
            continue

        batch_no = str(r.get('D') or '').strip()
        bin_code = str(r.get('E') or '').strip()
        cdate = excel_date(r.get('C'))

        entry = stock_by_code[code]
        entry['code'] = code
        if wh == '01':
            entry['wh01Qty'] += qty
            entry['batches01'] += 1
        elif wh == '60':
            entry['wh60Qty'] += qty
            entry['batches60'] += 1
        entry['totalQty'] = entry['wh01Qty'] + entry['wh60Qty']
        if bin_code:
            entry['bins'].add(bin_code)
        if cdate and cdate > entry['latestDate']:
            entry['latestDate'] = cdate
        if len(entry['batchList']) < 15:
            entry['batchList'].append({
                'wh': wh,
                'batch': batch_no,
                'bin': bin_code,
                'date': cdate,
                'qty': qty
            })

    # Format stock items list
    stock_items = []
    for code in sorted(stock_by_code.keys()):
        d = stock_by_code[code]
        meta = item_meta.get(code, {})
        feat = meta.get('feature', '')
        desc = meta.get('description', '')
        is_fg = code.startswith('8') and len(code) == 10
        if is_fg and not feat:
            feat = code[1:5]
        side = detect_side(code, desc)
        category = 'Finished Goods (Thành Phẩm Đầu 8)' if is_fg else (
            'Side Assembly (Bán Thành Phẩm)' if code.startswith('1220') else
            'Stamp Part (Linh Kiện Dập 34/38/39)' if code.startswith(('34', '36', '38', '39')) else
            'Hardware & Accessories (Phụ Kiện)'
        )
        stock_items.append({
            'code': code,
            'feature': feat or code[:4],
            'description': desc or f'Item {code}',
            'customerCode': meta.get('customerCode', ''),
            'category': category,
            'isFinishedGood': is_fg,
            'side': side,  # 'L', 'R', or 'N'
            'wh01Qty': d['wh01Qty'],
            'wh60Qty': d['wh60Qty'],
            'totalQty': d['totalQty'],
            'batches01': d['batches01'],
            'batches60': d['batches60'],
            'bins': sorted(list(d['bins'])),
            'latestDate': d['latestDate'],
            'batchList': d['batchList']
        })

    # 3. Build Left/Right (Trái/Phải) Paired Models for 3D Dashboard
    fg_pairs_map = defaultdict(dict)
    for item in stock_items:
        code = item['code']
        if item['isFinishedGood']:
            pair_key = code[:5] + 'X' + code[6:]
            model_id = code[1:5]
            side = 'L' if code[5] == '1' else ('R' if code[5] == '2' else 'N')
            fg_pairs_map[pair_key]['model'] = model_id
            fg_pairs_map[pair_key]['pairKey'] = pair_key
            fg_pairs_map[pair_key]['isFinishedGood'] = True
            fg_pairs_map[pair_key][side] = item

    # Also add explicit Left/Right semi-finished / hardware pairs in WH 01 & 60
    explicit_pairs = [
        ('1220XX0004', '1220', '1220190004', '1220200004', False),
        ('14217X0201', '1421', '1421710201', '1421720201', False),
        ('1427XX0001', '1427', '1427790001', '1427800001', False),
        ('14553X0001', '1455', '1455350001', '1455360001', False),
    ]
    code_lookup = {x['code']: x for x in stock_items}
    for pk, mid, lcode, rcode, is_fg in explicit_pairs:
        if lcode in code_lookup or rcode in code_lookup:
            fg_pairs_map[pk] = {
                'model': mid,
                'pairKey': pk,
                'isFinishedGood': is_fg,
                'L': code_lookup.get(lcode),
                'R': code_lookup.get(rcode)
            }

    lr_pairs = []
    for pk in sorted(fg_pairs_map.keys()):
        entry = fg_pairs_map[pk]
        l_item = entry.get('L') or {}
        r_item = entry.get('R') or {}
        l_01 = l_item.get('wh01Qty', 0.0)
        l_60 = l_item.get('wh60Qty', 0.0)
        l_tot = l_item.get('totalQty', 0.0)
        r_01 = r_item.get('wh01Qty', 0.0)
        r_60 = r_item.get('wh60Qty', 0.0)
        r_tot = r_item.get('totalQty', 0.0)
        diff = l_tot - r_tot
        matched_sets = min(l_tot, r_tot)
        wh_internal_imbalance = (l_01 != r_01) or (l_60 != r_60)
        desc = l_item.get('description') or r_item.get('description') or f"Model {entry['model']}"
        lr_pairs.append({
            'pairKey': pk,
            'model': entry['model'],
            'isFinishedGood': entry['isFinishedGood'],
            'description': desc,
            'leftCode': l_item.get('code', f"{pk[:5]}1{pk[6:]}"),
            'rightCode': r_item.get('code', f"{pk[:5]}2{pk[6:]}"),
            'leftWh01': l_01,
            'leftWh60': l_60,
            'leftTotal': l_tot,
            'rightWh01': r_01,
            'rightWh60': r_60,
            'rightTotal': r_tot,
            'combinedTotal': l_tot + r_tot,
            'matchedSets': matched_sets,
            'diffLR': diff,
            'isBalanced': diff == 0,
            'hasWarehouseCrossBalance': (diff == 0 and wh_internal_imbalance)
        })

    # 4. Parse Shipment Plans: NOV-full (Monthly), OCT-WK 42 (Weekly), Data (OPENING), booking follow
    def parse_plan_sheet(sname):
        rows = ship_sheets.get(sname, [])[1:]
        lines = []
        containers = []
        current_lines = []
        cont_idx = 1

        for r in rows:
            code = str(r.get('F') or '').strip()
            if code and code != 'None':
                qty = safe_float(r.get('N'), 0.0)
                line_obj = {
                    'id': f"{sname}-R{r.get('_row')}",
                    'sheet': sname,
                    'rowNumber': r.get('_row'),
                    'containerIndex': cont_idx,
                    'customer': str(r.get('A') or '').strip(),
                    'type': str(r.get('B') or '').strip(),
                    'pkl': str(r.get('C') or '').strip() if r.get('C') else '',
                    'so': str(r.get('D') or '').strip() if r.get('D') else '',
                    'po': str(r.get('E') or '').strip() if r.get('E') else '',
                    'itemCode': code,
                    'customerItemCode': str(r.get('G') or '').strip() if r.get('G') else '',
                    'feature': str(r.get('H') or '').strip() if r.get('H') else '',
                    'description': str(r.get('I') or '').strip() if r.get('I') else '',
                    'dimension': str(r.get('J') or '').strip() if r.get('J') and not str(r.get('J')).startswith('#') else '',
                    'crd': excel_date(r.get('K')),
                    'loadingDate': excel_date(r.get('L')),
                    'updatedLoadingDate': excel_date(r.get('M')),
                    'qty': qty,
                    'unitWeight': safe_float(r.get('O'), 0.0),
                    'pcsPerPkg': safe_float(r.get('P'), 0.0),
                    'pkgs': safe_float(r.get('Q'), 0.0),
                    'netWeight': safe_float(r.get('R'), 0.0),
                    'grossWeight': safe_float(r.get('S'), 0.0),
                    'cbm': safe_float(r.get('Y'), 0.0),
                    'unitPrice': safe_float(r.get('AA'), 0.0),
                    'totalAmount': safe_float(r.get('AB'), 0.0),
                    'side': detect_side(code, str(r.get('I') or ''))
                }
                lines.append(line_obj)
                current_lines.append(line_obj)
            else:
                # Subtotal / container boundary row
                if current_lines:
                    sub_qty = safe_float(r.get('N'), sum(x['qty'] for x in current_lines))
                    sub_pkg = safe_float(r.get('Q'), sum(x['pkgs'] for x in current_lines))
                    sub_nw = safe_float(r.get('R'), sum(x['netWeight'] for x in current_lines))
                    sub_gw = safe_float(r.get('S'), 0.0)
                    sub_cbm = safe_float(r.get('Y'), sum(x['cbm'] for x in current_lines))
                    sub_amt = safe_float(r.get('AB'), sum(x['totalAmount'] for x in current_lines))
                    pos = sorted(list({x['po'] for x in current_lines if x['po']}))
                    sos = sorted(list({x['so'] for x in current_lines if x['so']}))
                    custs = sorted(list({x['customer'] for x in current_lines if x['customer']}))
                    containers.append({
                        'containerIndex': cont_idx,
                        'sheet': sname,
                        'subtotalRow': r.get('_row'),
                        'customer': ', '.join(custs),
                        'pos': pos,
                        'sos': sos,
                        'pkl': current_lines[0].get('pkl', ''),
                        'loadingDate': current_lines[0].get('updatedLoadingDate') or current_lines[0].get('loadingDate'),
                        'lineCount': len(current_lines),
                        'totalQty': sub_qty,
                        'totalPkgs': sub_pkg,
                        'netWeight': sub_nw,
                        'grossWeight': sub_gw,
                        'totalCbm': sub_cbm,
                        'totalAmount': sub_amt
                    })
                    current_lines = []
                    cont_idx += 1
        return lines, containers

    nov_lines, nov_containers = parse_plan_sheet('NOV-full')
    oct_lines, oct_containers = parse_plan_sheet('OCT-WK 42')

    # Parse Data sheet (OPENING orders)
    opening_lines = []
    for r in ship_sheets.get('Data', [])[1:]:
        if str(r.get('R') or '').strip() == 'OPENING':
            code = str(r.get('D') or '').strip()
            if not code or code == 'None':
                continue
            qty = safe_float(r.get('O'), 0.0)
            opening_lines.append({
                'id': f"DATA-R{r.get('_row')}",
                'sheet': 'Data (OPENING)',
                'rowNumber': r.get('_row'),
                'containerIndex': 0,
                'customer': str(r.get('A') or '').strip(),
                'type': str(r.get('B') or '').strip(),
                'pkl': '',
                'so': '',
                'po': str(r.get('C') or '').strip(),
                'itemCode': code,
                'customerItemCode': '',
                'feature': str(r.get('E') or '').strip(),
                'description': str(r.get('F') or '').strip(),
                'dimension': '',
                'crd': excel_date(r.get('G')),
                'loadingDate': excel_date(r.get('K')),
                'updatedLoadingDate': excel_date(r.get('L') or r.get('N')),
                'qty': qty,
                'unitWeight': 0.0,
                'pcsPerPkg': 0.0,
                'pkgs': 0.0,
                'netWeight': 0.0,
                'grossWeight': 0.0,
                'cbm': 0.0,
                'unitPrice': safe_float(r.get('P'), 0.0),
                'totalAmount': safe_float(r.get('Q'), 0.0),
                'month': str(r.get('U') or '').strip(),
                'status': 'OPENING',
                'side': detect_side(code, str(r.get('F') or ''))
            })

    # Parse booking follow
    bookings = []
    for r in ship_sheets.get('booking follow', [])[1:]:
        if r.get('_row', 0) <= 3:
            continue
        cust = str(r.get('C') or '').strip()
        po = str(r.get('D') or '').strip()
        if not cust and not po:
            continue
        bookings.append({
            'rowNumber': r.get('_row'),
            'contNo': str(r.get('B') or '').strip() if r.get('B') else '',
            'customer': cust,
            'po': po,
            'grossWeight': safe_float(r.get('F'), 0.0),
            'bookingReqDate': excel_date(r.get('G')),
            'crd': excel_date(r.get('H')),
            'loadingDate': str(r.get('I') or '').strip() if r.get('I') else '',
            'vendorBooking': str(r.get('K') or '').strip() if r.get('K') else '',
            'bookingNo': str(r.get('L') or '').strip() if r.get('L') else '',
            'etd': excel_date(r.get('M')),
            'eta': excel_date(r.get('N')),
            'siCutOff': str(r.get('O') or '').strip() if r.get('O') else '',
            'cyCutOff': str(r.get('P') or '').strip() if r.get('P') else ''
        })

    output = {
        'generatedAt': datetime.now(timezone.utc).strftime('%Y-%m-%dT%H:%M:%SZ'),
        'sourceFiles': {
            'stockFile': 'Stock list.xlsx',
            'shipmentFile': 'NOV. SHIPMENT -OCT22.xlsx'
        },
        'auditSummary': {
            'totalRawStockRows': total_raw_batches,
            'wh01Rows': wh_summary['01']['rows'],
            'wh01UniqueCodes': len(wh_summary['01']['codes']),
            'wh01TotalQty': wh_summary['01']['qty'],
            'wh60Rows': wh_summary['60']['rows'],
            'wh60UniqueCodes': len(wh_summary['60']['codes']),
            'wh60TotalQty': wh_summary['60']['qty'],
            'combinedUniqueCodes': len(stock_items),
            'overlappingCodesCount': len(wh_summary['01']['codes'] & wh_summary['60']['codes']),
            'combinedTotalQty': wh_summary['01']['qty'] + wh_summary['60']['qty'],
            'finishedGoodsCodesCount': sum(1 for x in stock_items if x['isFinishedGood']),
            'finishedGoodsTotalQty': sum(x['totalQty'] for x in stock_items if x['isFinishedGood']),
            'finishedGoodsModelsCount': sum(1 for p in lr_pairs if p['isFinishedGood']),
            'balancedFgModelsCount': sum(1 for p in lr_pairs if p['isFinishedGood'] and p['isBalanced']),
            'imbalancedFgModelsCount': sum(1 for p in lr_pairs if p['isFinishedGood'] and not p['isBalanced']),
            'novFullLinesCount': len(nov_lines),
            'novFullPosCount': len({x['po'] for x in nov_lines if x['po']}),
            'novFullContainersCount': len(nov_containers),
            'octWk42LinesCount': len(oct_lines),
            'octWk42PosCount': len({x['po'] for x in oct_lines if x['po']}),
            'dataOpeningLinesCount': len(opening_lines)
        },
        'stockItems': stock_items,
        'lrPairs': lr_pairs,
        'plans': {
            'NOV-full': {
                'name': 'Kế Hoạch Xuất Hàng Tháng 11 (NOV-full)',
                'type': 'monthly',
                'lines': nov_lines,
                'containers': nov_containers
            },
            'OCT-WK 42': {
                'name': 'Kế Hoạch Xuất Hàng Tuần 42 (OCT-WK 42)',
                'type': 'weekly',
                'lines': oct_lines,
                'containers': oct_containers
            },
            'Data-OPENING': {
                'name': 'Tổng Hợp Đơn Hàng Đang Mở (Data - OPENING)',
                'type': 'master',
                'lines': opening_lines,
                'containers': []
            }
        },
        'bookings': bookings
    }

    out_dir = os.path.join(ROOT_DIR, 'public', 'data')
    os.makedirs(out_dir, exist_ok=True)
    out_file = os.path.join(out_dir, 'initial-data.json')
    with open(out_file, 'w', encoding='utf-8') as f:
        json.dump(output, f, ensure_ascii=False, indent=2)

    files_dir = os.path.join(ROOT_DIR, 'public', 'files')
    os.makedirs(files_dir, exist_ok=True)
    shutil.copy2(stock_path, os.path.join(files_dir, 'Stock list.xlsx'))
    shutil.copy2(ship_path, os.path.join(files_dir, 'NOV. SHIPMENT -OCT22.xlsx'))
    print('Successfully built public/data/initial-data.json:', json.dumps(output['auditSummary'], indent=2))

if __name__ == '__main__':
    build_dataset()
