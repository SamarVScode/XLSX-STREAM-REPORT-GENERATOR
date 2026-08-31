#!/usr/bin/env python3
import sys
import logging
from pathlib import Path
from collections import defaultdict
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

try:
    from python_calamine import CalamineWorkbook
    HAS_CALAMINE = True
except ImportError:
    HAS_CALAMINE = False

try:
    from config.dc_config import ALLOWED_DCS_SET_LOWER
except ImportError:
    ALLOWED_DCS_SET_LOWER = {'alg', 'ayp', 'deo', 'jhs', 'jnp', 'knp', 'mau', 'mrz', 'mth', 'mzn', 'rbr', 'spr', 'vns', 'all'}

log = logging.getLogger("ei_stream_server.tat_report")

def detect_hub_col(headers):
    priorities = ['source dc', 'source_dc', 'dc code', 'dc']
    lower = [str(h).strip().lower() if h is not None else '' for h in headers]
    for p in priorities:
        if p in lower:
            return lower.index(p)
    return 0

def detect_status_col(headers):
    lower = [str(h).strip().lower() if h is not None else '' for h in headers]
    if 'status_status' in lower:
        return lower.index('status_status')
    for idx, h in enumerate(lower):
        if 'status' in h:
            return idx
    return 1

def generate_tat_report(input_file: Path, output_file: Path):
    path = Path(input_file)
    log.info(f"Loading input workbook for SCM TAT Report: {path}")

    headers = []
    rows = []

    if HAS_CALAMINE:
        try:
            cal = CalamineWorkbook.from_path(str(path))
            sheet_map = {s.lower(): s for s in cal.sheet_names}
            target_sheet = sheet_map.get('data', cal.sheet_names[0])
            raw_python_rows = cal.get_sheet_by_name(target_sheet).to_python()
            if raw_python_rows:
                headers = raw_python_rows[0]
                rows = raw_python_rows[1:]
        except Exception:
            rows = []

    if not rows:
        wb_in = openpyxl.load_workbook(str(path), data_only=True, read_only=True)
        ws_data = wb_in['Data'] if 'Data' in wb_in.sheetnames else wb_in.active
        all_r = [list(r) for r in ws_data.iter_rows(values_only=True)]
        if all_r:
            headers = all_r[0]
            rows = all_r[1:]
        wb_in.close()

    hub_idx = detect_hub_col(headers)
    status_idx = detect_status_col(headers)

    filtered_rows = []
    for r in rows:
        hub_val = str(r[hub_idx]).strip().lower() if (len(r) > hub_idx and r[hub_idx] is not None) else ''
        if hub_val in ALLOWED_DCS_SET_LOWER or 'all' in ALLOWED_DCS_SET_LOWER:
            filtered_rows.append(r)

    hub_stats = defaultdict(lambda: {'Delivered': 0, 'OFD': 0, 'Undelivered': 0, 'Others': 0, 'Total': 0})
    for r in filtered_rows:
        hub = str(r[hub_idx]).strip().upper() if (len(r) > hub_idx and r[hub_idx] is not None) else 'UNKNOWN'
        raw_status = str(r[status_idx]).strip() if (len(r) > status_idx and r[status_idx] is not None) else ''
        status_clean = raw_status.upper()

        if 'DELIVERED' in status_clean and 'UNDELIVERED' not in status_clean:
            cat = 'Delivered'
        elif 'OFD' in status_clean or 'OUT FOR DELIVERY' in status_clean:
            cat = 'OFD'
        elif 'UNDELIVERED' in status_clean:
            cat = 'Undelivered'
        else:
            cat = 'Others'

        hub_stats[hub][cat] += 1
        hub_stats[hub]['Total'] += 1

    wb_out = openpyxl.Workbook()
    wb_out.remove(wb_out.active)

    ws_sum = wb_out.create_sheet('SCM tat performance summary')
    ws_sum.sheet_view.showGridLines = True

    f_title = Font(name="Calibri", size=13, bold=True, color="FFFFFF")
    f_header = Font(name="Calibri", size=10, bold=True, color="FFFFFF")
    f_data = Font(name="Calibri", size=10, color="1F2937")
    f_total = Font(name="Calibri", size=10, bold=True, color="1E1B4B")

    fill_title = PatternFill("solid", fgColor="1E1B4B")
    fill_header = PatternFill("solid", fgColor="312E81")
    fill_total = PatternFill("solid", fgColor="E0E7FF")

    align_center = Alignment(horizontal="center", vertical="center")
    align_left = Alignment(horizontal="left", vertical="center")

    thin_side = Side(style="thin", color="CBD5E1")
    bdr = Border(left=thin_side, right=thin_side, top=thin_side, bottom=thin_side)

    ws_sum.merge_cells('A1:G1')
    ws_sum['A1'] = "SCM TAT 24Hrs Performance Summary"
    ws_sum['A1'].font = f_title
    ws_sum['A1'].fill = fill_title
    ws_sum['A1'].alignment = align_center
    ws_sum.row_dimensions[1].height = 32

    for col_i in range(1, 8):
        c = ws_sum.cell(1, col_i)
        c.fill = fill_title
        c.border = bdr

    headers_sum = ['Delivery Hub', 'Delivered', 'OFD', 'Undelivered', 'Others', 'Total OFD', 'Performance %']
    for idx, h in enumerate(headers_sum, 1):
        cell = ws_sum.cell(2, idx, h)
        cell.font = f_header
        cell.fill = fill_header
        cell.alignment = align_center
        cell.border = bdr
    ws_sum.row_dimensions[2].height = 24

    sorted_hubs = sorted(hub_stats.keys())
    curr_r = 3
    tot_del = tot_ofd = tot_undel = tot_oth = tot_all = 0

    for hub in sorted_hubs:
        s = hub_stats[hub]
        deliv = s['Delivered']
        ofd = s['OFD']
        undel = s['Undelivered']
        oth = s['Others']
        total = s['Total']
        pct = (deliv / total * 100) if total > 0 else 0.0

        tot_del += deliv
        tot_ofd += ofd
        tot_undel += undel
        tot_oth += oth
        tot_all += total

        vals = [hub, deliv, ofd, undel, oth, total, round(pct, 1)]
        for c_idx, val in enumerate(vals, 1):
            cell = ws_sum.cell(curr_r, c_idx, val)
            cell.font = f_data
            cell.border = bdr
            cell.alignment = align_left if c_idx == 1 else align_center
            if c_idx == 7:
                cell.number_format = '0.0"%"'
            elif c_idx > 1:
                cell.number_format = '#,##0'

        ws_sum.row_dimensions[curr_r].height = 20
        curr_r += 1

    overall_pct = (tot_del / tot_all * 100) if tot_all > 0 else 0.0
    tot_vals = ['Total', tot_del, tot_ofd, tot_undel, tot_oth, tot_all, round(overall_pct, 1)]
    for c_idx, val in enumerate(tot_vals, 1):
        cell = ws_sum.cell(curr_r, c_idx, val)
        cell.font = f_total
        cell.fill = fill_total
        cell.border = bdr
        cell.alignment = align_left if c_idx == 1 else align_center
        if c_idx == 7:
            cell.number_format = '0.0"%"'
        elif c_idx > 1:
            cell.number_format = '#,##0'
    ws_sum.row_dimensions[curr_r].height = 22

    widths = [18, 14, 12, 14, 12, 14, 16]
    for i, w in enumerate(widths, 1):
        ws_sum.column_dimensions[get_column_letter(i)].width = w

    ws_raw = wb_out.create_sheet('SCM TAT raw data')
    ws_raw.append(headers)
    for r in filtered_rows:
        ws_raw.append(r)

    wb_out.save(str(output_file))
    try:
        wb_out.close()
    except Exception:
        pass
    log.info(f"Successfully generated SCM TAT Report: {output_file}")
