#!/usr/bin/env python3
"""
SCM TAT 24Hrs Performance Report Generator Module for ei_report_server
======================================================================
Reads 'Data' sheet from input Excel file, detects source DC and status columns,
filters rows matching allowed DCs from dc_config, computes completed vs pending counts,
and generates formatted output workbook:
  1. 'SCM tat performance summary' sheet
  2. 'SCM TAT raw data' sheet
"""

import sys
import logging
from pathlib import Path
from collections import defaultdict
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# Ensure current directory is in sys.path for dc_config import
CURRENT_DIR = Path(__file__).resolve().parent
if str(CURRENT_DIR) not in sys.path:
    sys.path.insert(0, str(CURRENT_DIR))
try:
    from config.dc_config import ALLOWED_DCS_SET_LOWER
except ImportError:
    try:
        from dc_config import ALLOWED_DCS_SET_LOWER
    except ImportError:
        ALLOWED_DCS_SET_LOWER = {'alg', 'ayp', 'deo', 'jhs', 'jnp', 'knp', 'mau', 'mrz', 'mth', 'mzn', 'rbr', 'spr', 'vns', 'all'}

log = logging.getLogger("ei_stream_server.tat_report")

def detect_hub_col(headers):
    priorities = ['source dc', 'source_dc', 'dc code', 'dc']
    lower = [str(h).strip().lower() if h is not None else '' for h in headers]
    for p in priorities:
        if p in lower:
            return lower.index(p)
    return 28  # Default fallback

def detect_status_col(headers):
    lower = [str(h).strip().lower() if h is not None else '' for h in headers]
    if 'status_status' in lower:
        return lower.index('status_status')
    for idx, h in enumerate(lower):
        if 'status' in h:
            return idx
    return 4  # Default fallback

def generate_tat_report(input_file: Path, output_file: Path):
    log.info(f"Loading input workbook for SCM TAT Report: {input_file}")
    wb_in = openpyxl.load_workbook(str(input_file), data_only=True)
    ws_data = wb_in['Data'] if 'Data' in wb_in.sheetnames else wb_in.active

    headers = [cell.value for cell in ws_data[1]]
    rows = []
    for row in ws_data.iter_rows(min_row=2, values_only=True):
        rows.append(list(row))
    wb_in.close()

    hub_idx = detect_hub_col(headers)
    status_idx = detect_status_col(headers)

    # Filter rows to allowed DCs
    filtered_rows = [
        r for r in rows
        if len(r) > hub_idx and str(r[hub_idx] or '').strip().lower() in ALLOWED_DCS_SET_LOWER
    ]

    stats_map = defaultdict(lambda: {'complete': 0, 'not_complete': 0})
    for row in filtered_rows:
        if len(row) <= max(hub_idx, status_idx):
            continue
        hub = str(row[hub_idx] or '').strip()
        status = str(row[status_idx] or '').strip().lower()
        if not hub or hub.lower() not in ALLOWED_DCS_SET_LOWER:
            continue
        if status in ('closed', 'task resolved', 'resolved', 'complete', 'completed'):
            stats_map[hub.upper()]['complete'] += 1
        else:
            stats_map[hub.upper()]['not_complete'] += 1

    sorted_hubs = sorted(stats_map.keys())
    summary_rows = []
    grand_complete = 0
    grand_not_complete = 0
    grand_total = 0

    for hub in sorted_hubs:
        s = stats_map[hub]
        total = s['complete'] + s['not_complete']
        complete_pct = s['complete'] / total if total > 0 else 0
        not_complete_pct = s['not_complete'] / total if total > 0 else 0
        grand_complete += s['complete']
        grand_not_complete += s['not_complete']
        grand_total += total
        summary_rows.append([hub, s['complete'], s['not_complete'], total, complete_pct, not_complete_pct])

    grand_complete_pct = grand_complete / grand_total if grand_total > 0 else 0
    grand_not_complete_pct = grand_not_complete / grand_total if grand_total > 0 else 0
    totals_row = ['TOTAL / SUMMARY', grand_complete, grand_not_complete, grand_total, grand_complete_pct, grand_not_complete_pct]

    # --- Styles ---
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
    green_fill = PatternFill(start_color='FFdcfce7', end_color='FFdcfce7', fill_type='solid')
    yellow_fill = PatternFill(start_color='FFfef9c3', end_color='FFfef9c3', fill_type='solid')
    red_fill = PatternFill(start_color='FFfee2e2', end_color='FFfee2e2', fill_type='solid')
    raw_header_fill = PatternFill(start_color='FF334155', end_color='FF334155', fill_type='solid')

    white_font = Font(bold=True, size=10, color='FFFFFFFF')
    green_font = Font(bold=True, size=9, color='FF166534')
    yellow_font = Font(bold=True, size=9, color='FF854d0e')
    red_font = Font(bold=True, size=9, color='FF991b1b')
    normal_font = Font(size=9)
    header_font = Font(bold=True, size=10, color='FFFFFFFF')
    raw_header_font = Font(bold=True, color='FFFFFFFF')

    center = Alignment(horizontal='center', vertical='center')

    wb_out = openpyxl.Workbook()
    ws_sum = wb_out.active
    ws_sum.title = 'Summary'
    ws_sum.sheet_view.showGridLines = False

    for r in range(1, len(summary_rows) + 10):
        for c in range(1, 8):
            ws_sum.cell(row=r, column=c).fill = white_fill

    date_str = 'SCM TAT 24Hrs Performance Report'
    ws_sum.cell(row=1, column=1, value=date_str)
    ws_sum.merge_cells(start_row=1, start_column=1, end_row=1, end_column=6)
    banner_cell = ws_sum.cell(row=1, column=1)
    banner_cell.fill = banner_fill
    banner_cell.font = white_font
    banner_cell.alignment = center
    ws_sum.row_dimensions[1].height = 22

    sum_headers = ['DC', 'Completed', 'Pending', 'Total', 'Done %', 'Pending %']
    for i, h in enumerate(sum_headers, 1):
        cell = ws_sum.cell(row=3, column=i, value=h)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = center
        cell.border = purple_border
    ws_sum.row_dimensions[3].height = 24

    for r_idx, srow in enumerate(summary_rows):
        row_num = 4 + r_idx
        is_alt = r_idx % 2 == 1
        fill = alt_fill if is_alt else white_fill

        for c_idx, val in enumerate(srow, 1):
            cell = ws_sum.cell(row=row_num, column=c_idx, value=val)
            cell.alignment = center
            cell.border = data_border
            cell.fill = fill
            cell.font = normal_font

        for c in [2, 3, 4]:
            ws_sum.cell(row=row_num, column=c).number_format = '#,##0'

        ws_sum.cell(row=row_num, column=5).number_format = '0.0%'
        ws_sum.cell(row=row_num, column=6).number_format = '0.0%'

        complete_pct = srow[4]
        pct_cell = ws_sum.cell(row=row_num, column=5)
        pct_cell.font = Font(bold=True, size=9)
        if complete_pct >= 0.90:
            pct_cell.fill = green_fill
            pct_cell.font = green_font
        elif complete_pct >= 0.75:
            pct_cell.fill = yellow_fill
            pct_cell.font = yellow_font
        else:
            pct_cell.fill = red_fill
            pct_cell.font = red_font

        ws_sum.row_dimensions[row_num].height = 20

    totals_row_num = 4 + len(summary_rows)
    for c_idx, val in enumerate(totals_row, 1):
        cell = ws_sum.cell(row=totals_row_num, column=c_idx, value=val)
        cell.alignment = center
        cell.border = purple_border
        cell.fill = header_fill
        cell.font = header_font

    ws_sum.cell(row=totals_row_num, column=2).number_format = '#,##0'
    ws_sum.cell(row=totals_row_num, column=3).number_format = '#,##0'
    ws_sum.cell(row=totals_row_num, column=4).number_format = '#,##0'
    ws_sum.cell(row=totals_row_num, column=5).number_format = '0.0%'
    ws_sum.cell(row=totals_row_num, column=6).number_format = '0.0%'

    for col in range(1, 7):
        max_len = len(str(sum_headers[col - 1]))
        for r in range(4, totals_row_num + 1):
            val = ws_sum.cell(row=r, column=col).value
            if val is not None:
                max_len = max(max_len, len(str(val)))
        ws_sum.column_dimensions[get_column_letter(col)].width = max_len + 4

    ws_raw = wb_out.create_sheet('Raw')
    for i, h in enumerate(headers, 1):
        cell = ws_raw.cell(row=1, column=i, value=h)
        cell.fill = raw_header_fill
        cell.font = raw_header_font
        cell.alignment = center

    for r_idx, row in enumerate(filtered_rows, 2):
        for c_idx, val in enumerate(row, 1):
            ws_raw.cell(row=r_idx, column=c_idx, value=val)

    for col in range(1, min(len(headers) + 1, 21)):
        max_len = len(str(headers[col - 1])) if col - 1 < len(headers) else 10
        for r in range(2, min(len(filtered_rows) + 1, 102)):
            val = ws_raw.cell(row=r, column=col).value
            if val is not None:
                max_len = max(max_len, len(str(val)))
        ws_raw.column_dimensions[get_column_letter(col)].width = min(max_len + 2, 30)

    wb_out.save(str(output_file))
    log.info(f"Successfully generated SCM TAT Report: {output_file}")
