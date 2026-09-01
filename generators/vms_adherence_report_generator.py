#!/usr/bin/env python3
"""
VMS Adherence Report Generator Module for ei_stream_server
==========================================================
Reads 'Raw' sheet from VMS Adherence Excel file, filters rows where Source DC is in allowed list,
computes summary stats by Source DC, and generates output workbook:
  1. Summary Sheet (VMS Adherence Summary Table with % and color highlights)
  2. Raw Sheet (Full filtered dataset)

Uses Centralized Zero-Memory Streaming Engine (core.stream_engine):
- O(1) Memory Footprint (< 35MB RAM)
- Fast Rust / openpyxl stream reader + direct disk XML streaming
"""

import sys
import logging
from pathlib import Path
from collections import defaultdict
import openpyxl
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# Ensure server root is in sys.path
SERVER_ROOT = Path(__file__).resolve().parent.parent
if str(SERVER_ROOT) not in sys.path:
    sys.path.insert(0, str(SERVER_ROOT))

try:
    from config.dc_config import ALLOWED_DCS_SET_LOWER
except ImportError:
    from dc_config import ALLOWED_DCS_SET_LOWER

from core.stream_engine import (
    XmlSheetWriter,
    assemble_stream_workbook,
    stream_sheet_rows,
    inspect_spreadsheet_headers
)

log = logging.getLogger("ei_stream_server.vms_adherence_report")


def find_col(headers, names):
    lower = [str(h).strip().lower() if h is not None else '' for h in headers]
    for name in names:
        if name in lower:
            return lower.index(name)
    raise Exception(f'Column not found: {names}')


def generate_vms_adherence_report(input_file: Path, output_file: Path):
    input_path = Path(input_file)
    output_path = Path(output_file)
    log.info(f"Loading input workbook for VMS Adherence Report (Stream Mode): {input_path}")

    # Inspect headers
    headers = inspect_spreadsheet_headers(input_path)
    if not headers:
        # Fallback via first row iterator
        row_iter = stream_sheet_rows(input_path, start_row=1)
        headers = next(row_iter, [])

    if not headers:
        raise ValueError(f"Input file is empty: {input_path}")

    source_dc_idx = find_col(headers, ['source_dc', 'source dc', 'sourcedc', 'dc'])
    vms_status_idx = find_col(headers, ['vms status', 'vms_status', 'vmsstatus', 'status'])

    stats = defaultdict(lambda: {'done': 0, 'not_done': 0})
    total_filtered = 0

    # 1. Stream Raw data directly to disk XML (0 MB RAM)
    raw_writer = XmlSheetWriter("Raw", headers)
    with raw_writer:
        for row in stream_sheet_rows(input_path, start_row=2):
            if not row or len(row) <= source_dc_idx:
                continue
            raw_dc = row[source_dc_idx]
            if raw_dc is None:
                continue
            dc_clean = str(raw_dc).strip().lower()

            if dc_clean in ALLOWED_DCS_SET_LOWER:
                total_filtered += 1
                status = str(row[vms_status_idx] or '').strip().lower() if len(row) > vms_status_idx else ''
                if status == 'done':
                    stats[dc_clean]['done'] += 1
                else:
                    stats[dc_clean]['not_done'] += 1

                raw_writer.write_row(row)

    log.info(f"Filtered {total_filtered} matching rows.")

    # 2. Build styled Summary sheet in memory (~40 KB RAM)
    sorted_dcs = sorted(stats.keys())
    summary_rows = []
    for dc in sorted_dcs:
        d = stats[dc]
        total = d['done'] + d['not_done']
        done_pct = d['done'] / total if total > 0 else 0
        summary_rows.append([dc.upper(), total, d['done'], d['not_done'], done_pct])

    wb_out = Workbook()
    ws_sum = wb_out.active
    ws_sum.title = 'Summary'
    ws_sum.sheet_view.showGridLines = False

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

    white_font = Font(bold=True, size=10, color='FFFFFFFF')
    green_font = Font(bold=True, size=9, color='FF166534')
    yellow_font = Font(bold=True, size=9, color='FF854d0e')
    red_font = Font(bold=True, size=9, color='FF991b1b')
    normal_font = Font(size=9)
    header_font = Font(bold=True, size=10, color='FFFFFFFF')
    center = Alignment(horizontal='center', vertical='center')

    # Fill white background
    for r in range(1, len(summary_rows) + 10):
        for c in range(1, 8):
            ws_sum.cell(row=r, column=c).fill = white_fill

    # Banner
    ws_sum.cell(row=1, column=1, value='VMS Adherence Summary')
    ws_sum.merge_cells(start_row=1, start_column=1, end_row=1, end_column=5)
    banner_cell = ws_sum.cell(row=1, column=1)
    banner_cell.fill = banner_fill
    banner_cell.font = white_font
    banner_cell.alignment = center
    ws_sum.row_dimensions[1].height = 22

    # Headers
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
        is_alt = (r_idx % 2 == 1)
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

        # Done % color
        done_pct = srow[4]
        pct_cell = ws_sum.cell(row=row_num, column=5)
        pct_cell.font = Font(bold=True, size=9)
        if done_pct > 0.85:
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
        ws_sum.column_dimensions[get_column_letter(col)].width = max_len + 4

    # 3. Assemble final .xlsx in 1 line
    assemble_stream_workbook(wb_out, [raw_writer], output_path)
    log.info(f"Successfully generated VMS Adherence Report: {output_file.name} ({total_filtered} rows)")
