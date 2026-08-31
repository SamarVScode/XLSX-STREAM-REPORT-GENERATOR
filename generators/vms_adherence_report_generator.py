#!/usr/bin/env python3
"""
VMS Adherence Stream Report Generator Module for ei_stream_server
=================================================================
Streams rows from VMS Adherence spreadsheet, computes done/not done metrics per DC in-flight,
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

log = logging.getLogger("ei_stream_server.vms_adherence_report")

def find_col(headers, names, default=0):
    lower = [str(h).strip().lower().replace('_', ' ') if h is not None else '' for h in headers]
    for name in names:
        n_clean = name.lower().replace('_', ' ')
        if n_clean in lower:
            return lower.index(n_clean)
    for idx, h in enumerate(lower):
        for name in names:
            n_clean = name.lower().replace('_', ' ')
            if n_clean in h:
                return idx
    return default

def generate_vms_adherence_report(input_file: Path, output_file: Path):
    path = Path(input_file)
    log.info(f"Stream generating VMS Adherence Report: {path}")

    sheet_names = get_sheet_names(path)
    target_sheet = 'Raw' if 'Raw' in sheet_names else sheet_names[0]

    rows_iter = stream_sheet_rows(path, sheet_name=target_sheet)
    try:
        raw_headers = next(rows_iter)
    except StopIteration:
        raise ValueError(f"Sheet '{target_sheet}' is empty.")

    headers = [str(h).strip() if h is not None else '' for h in raw_headers]
    source_dc_idx = find_col(headers, ['source_dc', 'source dc', 'sourcedc', 'dc', 'sdc'], default=0)
    vms_status_idx = find_col(headers, ['vms status', 'vms_status', 'vmsstatus', 'status'], default=1)

    filtered_rows = []
    stats = defaultdict(lambda: {'done': 0, 'not_done': 0})

    for row in rows_iter:
        if not row or len(row) <= max(source_dc_idx, vms_status_idx):
            continue
        dc_key = str(row[source_dc_idx] or '').strip().lower()
        if dc_key in ALLOWED_DCS_SET_LOWER:
            filtered_rows.append(row)
            status = str(row[vms_status_idx] or '').strip().lower()
            if status in ('done', 'completed', 'adherence'):
                stats[dc_key]['done'] += 1
            else:
                stats[dc_key]['not_done'] += 1

    log.info(f"Streamed and filtered {len(filtered_rows)} matching VMS rows.")

    sorted_dcs = sorted(stats.keys())
    summary_rows = []
    grand_done = 0
    grand_not_done = 0

    for dc in sorted_dcs:
        d = stats[dc]
        total = d['done'] + d['not_done']
        done_pct = d['done'] / total if total > 0 else 0
        grand_done += d['done']
        grand_not_done += d['not_done']
        summary_rows.append([dc.upper(), total, d['done'], d['not_done'], done_pct])

    grand_total = grand_done + grand_not_done
    grand_done_pct = grand_done / grand_total if grand_total > 0 else 0

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
    green_fill = PatternFill(start_color='FFdcfce7', end_color='FFdcfce7', fill_type='solid')
    yellow_fill = PatternFill(start_color='FFfef9c3', end_color='FFfef9c3', fill_type='solid')
    red_fill = PatternFill(start_color='FFfee2e2', end_color='FFfee2e2', fill_type='solid')

    banner_font = Font(name='Calibri', size=13, bold=True, color='FFFFFFFF')
    header_font = Font(name='Calibri', size=11, bold=True, color='FFFFFFFF')
    data_font = Font(name='Calibri', size=11, color='FF1e1b4b')
    dc_font = Font(name='Calibri', size=11, bold=True, color='FF1e1b4b')
    green_font = Font(name='Calibri', size=11, bold=True, color='FF15803d')
    yellow_font = Font(name='Calibri', size=11, bold=True, color='FFa16207')
    red_font = Font(name='Calibri', size=11, bold=True, color='FFb91c1c')

    center_align = Alignment(horizontal='center', vertical='center')
    right_align = Alignment(horizontal='right', vertical='center')

    wb = openpyxl.Workbook()

    # --- Summary Sheet ---
    ws_sum = wb.active
    ws_sum.title = 'Summary'
    ws_sum.views.sheetView[0].showGridLines = True

    # Row 1: Banner
    ws_sum.merge_cells('B1:F1')
    ws_sum.row_dimensions[1].height = 36
    banner_cell = ws_sum.cell(row=1, column=2, value='VMS Adherence Summary')
    banner_cell.font = banner_font
    banner_cell.fill = banner_fill
    banner_cell.alignment = center_align

    for col in range(2, 7):
        c = ws_sum.cell(row=1, column=col)
        c.fill = banner_fill
        c.border = purple_border

    # Row 2: Headers
    headers_summary = ['Source DC', 'Total Shipments', 'Done Count', 'Not Done Count', 'Done%']
    ws_sum.row_dimensions[2].height = 24
    for c_idx, h_text in enumerate(headers_summary, start=2):
        cell = ws_sum.cell(row=2, column=c_idx, value=h_text)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = center_align
        cell.border = purple_border

    # Data Rows
    current_r = 3
    for row_idx, r_vals in enumerate(summary_rows):
        ws_sum.row_dimensions[current_r].height = 20
        is_alt = (row_idx % 2 == 1)
        row_fill = alt_fill if is_alt else white_fill
        pct_val = r_vals[4]

        if pct_val > 0.85:
            pct_fill = green_fill
            pct_f = green_font
        elif pct_val >= 0.65:
            pct_fill = yellow_fill
            pct_f = yellow_font
        else:
            pct_fill = red_fill
            pct_f = red_font

        for c_offset, val in enumerate(r_vals):
            col_num = 2 + c_offset
            cell = ws_sum.cell(row=current_r, column=col_num, value=val)
            cell.border = data_border
            if c_offset == 0:
                cell.font = dc_font
                cell.fill = row_fill
                cell.alignment = center_align
            elif c_offset in (1, 2, 3):
                cell.font = data_font
                cell.fill = row_fill
                cell.alignment = right_align
                cell.number_format = '#,##0'
            elif c_offset == 4:
                cell.font = pct_f
                cell.fill = pct_fill
                cell.alignment = center_align
                cell.number_format = '0.0%'
        current_r += 1

    # Total Row
    ws_sum.row_dimensions[current_r].height = 22
    total_vals = ['TOTAL / SUMMARY', grand_total, grand_done, grand_not_done, grand_done_pct]
    if grand_done_pct > 0.85:
        tot_pct_fill = green_fill
        tot_pct_font = green_font
    elif grand_done_pct >= 0.65:
        tot_pct_fill = yellow_fill
        tot_pct_font = yellow_font
    else:
        tot_pct_fill = red_fill
        tot_pct_font = red_font

    for c_offset, val in enumerate(total_vals):
        col_num = 2 + c_offset
        cell = ws_sum.cell(row=current_r, column=col_num, value=val)
        cell.border = purple_border
        if c_offset == 0:
            cell.font = Font(name='Calibri', size=11, bold=True, color='FFFFFFFF')
            cell.fill = banner_fill
            cell.alignment = center_align
        elif c_offset in (1, 2, 3):
            cell.font = Font(name='Calibri', size=11, bold=True, color='FFFFFFFF')
            cell.fill = banner_fill
            cell.alignment = right_align
            cell.number_format = '#,##0'
        elif c_offset == 4:
            cell.font = tot_pct_font
            cell.fill = tot_pct_fill
            cell.alignment = center_align
            cell.number_format = '0.0%'

    # Column Widths
    ws_sum.column_dimensions['A'].width = 3
    col_widths = {'B': 18, 'C': 18, 'D': 16, 'E': 18, 'F': 14}
    for col_l, w in col_widths.items():
        ws_sum.column_dimensions[col_l].width = w

    # --- Raw Sheet ---
    ws_raw = wb.create_sheet(title='Raw')
    ws_raw.views.sheetView[0].showGridLines = True
    ws_raw.row_dimensions[1].height = 24

    for c_idx, h in enumerate(headers, start=1):
        cell = ws_raw.cell(row=1, column=c_idx, value=h)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = center_align
        cell.border = purple_border

    for r_idx, row_vals in enumerate(filtered_rows, start=2):
        ws_raw.row_dimensions[r_idx].height = 19
        use_alt = (r_idx % 2 == 1)
        r_fill = alt_fill if use_alt else white_fill
        for c_idx, val in enumerate(row_vals, start=1):
            cell = ws_raw.cell(row=r_idx, column=c_idx, value=val)
            cell.font = data_font
            cell.fill = r_fill
            cell.border = data_border
            if isinstance(val, (int, float)):
                cell.alignment = right_align
            else:
                cell.alignment = Alignment(horizontal='left', vertical='center')

    for col in ws_raw.columns:
        col_letter = get_column_letter(col[0].column)
        max_len = 0
        for cell in col:
            val_str = str(cell.value or '')
            max_len = max(max_len, len(val_str))
        ws_raw.column_dimensions[col_letter].width = min(max(max_len + 3, 12), 40)

    wb.save(output_file)
    log.info(f"Successfully generated VMS Adherence Report: {output_file}")
    return str(output_file)
