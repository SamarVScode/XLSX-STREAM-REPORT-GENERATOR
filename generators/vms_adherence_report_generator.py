#!/usr/bin/env python3
"""
VMS Adherence Report Generator Module for ei_stream_server
==========================================================
Reads 'Raw' sheet from VMS Adherence Excel file, filters rows where Source DC is in allowed list,
computes summary stats by Source DC, and generates output workbook:
  1. Summary Sheet (VMS Adherence Summary Table with % and color highlights)
  2. Raw Sheet (Full filtered dataset)
"""

import sys
import re
import logging
from pathlib import Path
from collections import defaultdict
import openpyxl
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# Ensure current directory is in sys.path for dc_config import
CURRENT_DIR = Path(__file__).resolve().parent
if str(CURRENT_DIR) not in sys.path:
    sys.path.insert(0, str(CURRENT_DIR))

try:
    from config.dc_config import ALLOWED_SOURCE_DCS, ALLOWED_DCS_SET, ALLOWED_DCS_SET_LOWER
except ImportError:
    from dc_config import ALLOWED_SOURCE_DCS, ALLOWED_DCS_SET, ALLOWED_DCS_SET_LOWER

log = logging.getLogger("ei_stream_server.vms_adherence_report")


def _clean_key(name) -> str:
    """Normalize string by removing all whitespace, punctuation, and lowercasing."""
    if name is None:
        return ""
    return re.sub(r'[^a-z0-9]', '', str(name).lower())


def find_col_index(headers, candidates, fallback_keywords=None, default=0) -> int:
    """Robust column detection supporting exact matches, cleaned keys, and substring fallbacks."""
    if not headers:
        return default

    clean_map = {_clean_key(h): idx for idx, h in enumerate(headers) if h is not None}
    raw_map = {str(h).strip().lower(): idx for idx, h in enumerate(headers) if h is not None}

    # 1. Exact cleaned match
    for cand in candidates:
        c_clean = _clean_key(cand)
        if c_clean in clean_map:
            return clean_map[c_clean]
        if cand.strip().lower() in raw_map:
            return raw_map[cand.strip().lower()]

    # 2. Keyword / Substring fallback
    if fallback_keywords:
        for kw in fallback_keywords:
            kw_clean = _clean_key(kw)
            for k, idx in clean_map.items():
                if kw_clean in k:
                    return idx

    return default


def normalize_vms_status(val) -> str:
    """Normalizes VMS adherence status to 'done' or 'not_done'."""
    if val is None:
        return "not_done"
    s = str(val).strip().lower()
    if s in ('done', 'completed', 'yes', 'y', '1', '1.0', 'true', 'adhered', 'success'):
        return "done"
    return "not_done"


def generate_vms_adherence_report(input_file: Path, output_file: Path):
    input_path = Path(input_file)
    output_path = Path(output_file)
    log.info(f"Loading input workbook for VMS Adherence Report: {input_path}")

    headers = []
    rows = []

    try:
        from python_calamine import CalamineWorkbook
        calamine_wb = CalamineWorkbook.from_path(str(input_path))
        
        # Detect sheet name case-insensitively
        sheet_map = {s.strip().lower(): s for s in calamine_wb.sheet_names}
        target_sheet = None
        for candidate in ['raw', 'raw_data', 'data', 'vms', 'vms_data', 'sheet1']:
            if candidate in sheet_map:
                target_sheet = sheet_map[candidate]
                break
        if not target_sheet and calamine_wb.sheet_names:
            target_sheet = calamine_wb.sheet_names[0]

        ws_raw = calamine_wb.get_sheet_by_name(target_sheet)
        raw_python_rows = ws_raw.to_python()
        if raw_python_rows:
            headers = raw_python_rows[0]
            rows = raw_python_rows[1:]
    except Exception as e:
        log.warning(f"Calamine read failed: {e}. Falling back to openpyxl.")
        wb_in = openpyxl.load_workbook(str(input_path), data_only=True, read_only=True)
        sheet_map = {s.strip().lower(): s for s in wb_in.sheetnames}
        target_sheet = None
        for candidate in ['raw', 'raw_data', 'data', 'vms', 'vms_data', 'sheet1']:
            if candidate in sheet_map:
                target_sheet = sheet_map[candidate]
                break
        if not target_sheet and wb_in.sheetnames:
            target_sheet = wb_in.sheetnames[0]

        ws_raw = wb_in[target_sheet]
        row_iter = ws_raw.iter_rows(values_only=True)
        first_row = next(row_iter, None)
        if first_row:
            headers = list(first_row)
            rows = [list(r) for r in row_iter if any(v is not None for v in r)]
        wb_in.close()

    if not headers:
        raise ValueError(f"Input file is empty or has no header row: {input_path}")

    # Column resolution
    source_dc_idx = find_col_index(
        headers,
        candidates=['source_dc', 'source dc', 'sourcedc', 'dc', 'sourcedccode', 'origin_dc', 'hub'],
        fallback_keywords=['sourcedc', 'dc', 'hub', 'origin'],
        default=0
    )

    vms_status_idx = find_col_index(
        headers,
        candidates=['vms status', 'vms_status', 'vmsstatus', 'status', 'vms', 'adherence_status', 'adherence'],
        fallback_keywords=['vms', 'adherence', 'status'],
        default=1 if len(headers) > 1 else 0
    )

    log.info(f"Resolved columns: Source_DC={source_dc_idx} ('{headers[source_dc_idx]}'), "
             f"VMS_Status={vms_status_idx} ('{headers[vms_status_idx]}')")

    # Filter rows to allowed DCs
    filtered_rows = []
    stats = defaultdict(lambda: {'done': 0, 'not_done': 0})

    for r in rows:
        if not r or len(r) <= source_dc_idx:
            continue
        raw_dc = r[source_dc_idx]
        if raw_dc is None:
            continue
        dc_clean = str(raw_dc).strip().lower()
        if dc_clean in ALLOWED_DCS_SET_LOWER:
            filtered_rows.append(r)
            dc_upper = dc_clean.upper()
            
            status_val = r[vms_status_idx] if len(r) > vms_status_idx else None
            norm_status = normalize_vms_status(status_val)
            if norm_status == 'done':
                stats[dc_upper]['done'] += 1
            else:
                stats[dc_upper]['not_done'] += 1

    log.info(f"Filtered {len(filtered_rows)} rows matching allowed DCs")

    # Build summary rows following ALLOWED_SOURCE_DCS order
    summary_rows = []
    grand_done = 0
    grand_not_done = 0

    for dc in ALLOWED_SOURCE_DCS:
        dc_u = dc.upper()
        if dc_u == 'ALL':
            continue
        d = stats.get(dc_u, {'done': 0, 'not_done': 0})
        total = d['done'] + d['not_done']
        if total > 0:
            done_pct = d['done'] / total
            grand_done += d['done']
            grand_not_done += d['not_done']
            summary_rows.append([dc_u, total, d['done'], d['not_done'], done_pct])

    # Also include any other active DCs in stats
    for dc_u, d in sorted(stats.items()):
        if dc_u not in ALLOWED_DCS_SET:
            total = d['done'] + d['not_done']
            if total > 0:
                done_pct = d['done'] / total
                grand_done += d['done']
                grand_not_done += d['not_done']
                summary_rows.append([dc_u, total, d['done'], d['not_done'], done_pct])

    grand_total = grand_done + grand_not_done
    grand_done_pct = grand_done / grand_total if grand_total > 0 else 0

    # Total row
    total_row = ['Total', grand_total, grand_done, grand_not_done, grand_done_pct]
    summary_rows.append(total_row)

    # Styles
    purple_border_color = 'FF3b0764'
    purple_border = Border(
        left=Side(style='thin', color=purple_border_color),
        right=Side(style='thin', color=purple_border_color),
        top=Side(style='thin', color=purple_border_color),
        bottom=Side(style='thin', color=purple_border_color)
    )
    data_border_color = 'FFddd6fe'
    data_border = Border(
        left=Side(style='thin', color=data_border_color),
        right=Side(style='thin', color=data_border_color),
        top=Side(style='thin', color=data_border_color),
        bottom=Side(style='thin', color=data_border_color)
    )

    banner_fill = PatternFill(start_color='FF4c1d95', end_color='FF4c1d95', fill_type='solid')
    header_fill = PatternFill(start_color='FF6d28d9', end_color='FF6d28d9', fill_type='solid')
    white_fill = PatternFill(start_color='FFFFFFFF', end_color='FFFFFFFF', fill_type='solid')
    alt_fill = PatternFill(start_color='FFF5f3ff', end_color='FFF5f3ff', fill_type='solid')
    total_fill = PatternFill(start_color='FFE0E7FF', end_color='FFE0E7FF', fill_type='solid')
    green_fill = PatternFill(start_color='FFdcfce7', end_color='FFdcfce7', fill_type='solid')
    yellow_fill = PatternFill(start_color='FFfef9c3', end_color='FFfef9c3', fill_type='solid')
    red_fill = PatternFill(start_color='FFfee2e2', end_color='FFfee2e2', fill_type='solid')
    raw_header_fill = PatternFill(start_color='FF334155', end_color='FF334155', fill_type='solid')

    white_font = Font(name='Calibri', bold=True, size=11, color='FFFFFFFF')
    total_font = Font(name='Calibri', bold=True, size=10, color='FF1E1B4B')
    green_font = Font(name='Calibri', bold=True, size=9, color='FF166534')
    yellow_font = Font(name='Calibri', bold=True, size=9, color='FF854d0e')
    red_font = Font(name='Calibri', bold=True, size=9, color='FF991b1b')
    normal_font = Font(name='Calibri', size=9)
    header_font = Font(name='Calibri', bold=True, size=10, color='FFFFFFFF')
    raw_header_font = Font(name='Calibri', bold=True, color='FFFFFFFF')

    center = Alignment(horizontal='center', vertical='center')

    wb_out = Workbook()

    # ===== Summary sheet =====
    ws_sum = wb_out.active
    ws_sum.title = 'Summary'
    ws_sum.sheet_view.showGridLines = False

    # Row 1: Banner
    ws_sum.cell(row=1, column=1, value='VMS Adherence Summary')
    ws_sum.merge_cells(start_row=1, start_column=1, end_row=1, end_column=5)
    banner_cell = ws_sum.cell(row=1, column=1)
    banner_cell.fill = banner_fill
    banner_cell.font = white_font
    banner_cell.alignment = center
    ws_sum.row_dimensions[1].height = 24

    for c in range(1, 6):
        ws_sum.cell(row=1, column=c).fill = banner_fill
        ws_sum.cell(row=1, column=c).border = purple_border

    # Row 3: Headers
    sum_headers = ['Source DC', 'Total', 'VMS Done', 'VMS Not Done', 'Done %']
    for i, h in enumerate(sum_headers, 1):
        cell = ws_sum.cell(row=3, column=i, value=h)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = center
        cell.border = purple_border
    ws_sum.row_dimensions[3].height = 24

    # Data rows
    for r_idx, srow in enumerate(summary_rows):
        row_num = 4 + r_idx
        is_total_row = (srow[0] == 'Total')
        is_alt = (r_idx % 2 == 1)

        if is_total_row:
            fill = total_fill
            row_font = total_font
        else:
            fill = alt_fill if is_alt else white_fill
            row_font = normal_font

        for c_idx, val in enumerate(srow, 1):
            cell = ws_sum.cell(row=row_num, column=c_idx, value=val)
            cell.alignment = center
            cell.border = data_border
            cell.fill = fill
            cell.font = row_font

        for c in [2, 3, 4]:
            ws_sum.cell(row=row_num, column=c).number_format = '#,##0'

        ws_sum.cell(row=row_num, column=5).number_format = '0.0%'

        # Done % highlight
        done_pct = srow[4]
        pct_cell = ws_sum.cell(row=row_num, column=5)
        if is_total_row:
            pct_cell.fill = total_fill
            pct_cell.font = total_font
        else:
            pct_cell.font = Font(name='Calibri', bold=True, size=9)
            if done_pct >= 0.85:
                pct_cell.fill = green_fill
                pct_cell.font = green_font
            elif done_pct >= 0.65:
                pct_cell.fill = yellow_fill
                pct_cell.font = yellow_font
            else:
                pct_cell.fill = red_fill
                pct_cell.font = red_font

        ws_sum.row_dimensions[row_num].height = 20

    # Auto-fit columns
    for col in range(1, 6):
        max_len = len(str(sum_headers[col - 1]))
        for r in range(4, 4 + len(summary_rows)):
            val = ws_sum.cell(row=r, column=col).value
            if val is not None:
                max_len = max(max_len, len(str(val)))
        ws_sum.column_dimensions[get_column_letter(col)].width = max(max_len + 4, 14)

    # ===== Raw Data sheet =====
    ws_raw_out = wb_out.create_sheet('Raw')

    for i, h in enumerate(headers, 1):
        cell = ws_raw_out.cell(row=1, column=i, value=h)
        cell.fill = raw_header_fill
        cell.font = raw_header_font
        cell.alignment = center

    for r_idx, row in enumerate(filtered_rows, 2):
        for c_idx, val in enumerate(row, 1):
            ws_raw_out.cell(row=r_idx, column=c_idx, value=val)

    for col in range(1, min(len(headers) + 1, 21)):
        max_len = len(str(headers[col - 1])) if col - 1 < len(headers) else 10
        for r in range(2, min(len(filtered_rows) + 1, 102)):
            val = ws_raw_out.cell(row=r, column=col).value
            if val is not None:
                max_len = max(max_len, len(str(val)))
        ws_raw_out.column_dimensions[get_column_letter(col)].width = min(max_len + 2, 30)

    output_path.parent.mkdir(parents=True, exist_ok=True)
    wb_out.save(str(output_path))
    log.info(f"Successfully generated VMS Adherence Report: {output_file.name} ({len(filtered_rows)} rows)")


if __name__ == "__main__":
    if len(sys.argv) > 2:
        generate_vms_adherence_report(Path(sys.argv[1]), Path(sys.argv[2]))
    else:
        print("Usage: python vms_adherence_report_generator.py <input_xlsx> <output_xlsx>")
