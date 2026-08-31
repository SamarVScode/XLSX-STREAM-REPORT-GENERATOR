#!/usr/bin/env python3
"""
SCM TAT Stream Report Generator Module for ei_stream_server
===========================================================
Streams rows from input spreadsheet, computes completed vs pending counts per DC in-flight,
and generates formatted output workbook.
"""

import sys
import logging
from pathlib import Path
from collections import defaultdict
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

from config.dc_config import ALLOWED_DCS_SET_LOWER
from core.stream_engine import stream_sheet_rows, get_sheet_names

log = logging.getLogger("ei_stream_server.tat_report")

def detect_hub_col(headers):
    priorities = ['source dc', 'source_dc', 'dc code', 'dc']
    lower = [str(h).strip().lower() if h is not None else '' for h in headers]
    for p in priorities:
        if p in lower:
            return lower.index(p)
    return 28

def detect_status_col(headers):
    lower = [str(h).strip().lower() if h is not None else '' for h in headers]
    if 'status_status' in lower:
        return lower.index('status_status')
    for idx, h in enumerate(lower):
        if 'status' in h:
            return idx
    return 4

def generate_tat_report(input_file: Path, output_file: Path):
    path = Path(input_file)
    log.info(f"Stream generating SCM TAT Report: {path}")

    sheet_names = get_sheet_names(path)
    target_sheet = 'Data' if 'Data' in sheet_names else sheet_names[0]

    rows_iter = stream_sheet_rows(path, sheet_name=target_sheet)
    try:
        raw_headers = next(rows_iter)
    except StopIteration:
        raise ValueError(f"Sheet '{target_sheet}' is empty.")

    headers = [str(h).strip() if h is not None else '' for h in raw_headers]
    hub_idx = detect_hub_col(headers)
    status_idx = detect_status_col(headers)

    filtered_rows = []
    stats_map = defaultdict(lambda: {'complete': 0, 'not_complete': 0})

    for row in rows_iter:
        if not row or len(row) <= max(hub_idx, status_idx):
            continue
        hub = str(row[hub_idx] or '').strip()
        if not hub or hub.lower() not in ALLOWED_DCS_SET_LOWER:
            continue

        filtered_rows.append(row)
        status = str(row[status_idx] or '').strip().lower()
        if status in ('closed', 'task resolved', 'resolved', 'complete', 'completed'):
            stats_map[hub.upper()]['complete'] += 1
        else:
            stats_map[hub.upper()]['not_complete'] += 1

    log.info(f"Streamed and filtered {len(filtered_rows)} matching TAT rows.")

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
    return str(output_file)
