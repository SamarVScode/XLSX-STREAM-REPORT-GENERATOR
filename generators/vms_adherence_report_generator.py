#!/usr/bin/env python3
"""
VMS Adherence Report Generator Module for ei_report_server
==========================================================
Reads 'Raw' sheet from VMS Adherence Excel file, filters rows where Source DC is in allowed list,
computes summary stats by Source DC, and generates output workbook:
  1. VMS Adherence Summary Sheet
  2. VMS Adherence Raw Data Sheet
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
    from dc_config import ALLOWED_DCS_SET_LOWER

log = logging.getLogger("ei_stream_server.vms_adherence_report")


def find_col(headers, names):
    lower = [str(h).strip().lower() if h is not None else '' for h in headers]
    for name in names:
        if name in lower:
            return lower.index(name)
    raise Exception(f'Column not found: {names}')


def generate_vms_adherence_report(input_file: Path, output_file: Path):
    log.info(f"Loading input workbook for VMS Adherence Report: {input_file}")
    try:
        from python_calamine import CalamineWorkbook
        calamine_wb = CalamineWorkbook.from_path(str(input_file))
        ws_raw = calamine_wb.get_sheet_by_name('Raw') if 'Raw' in calamine_wb.sheet_names else calamine_wb.get_sheet_by_index(0)
        raw_python_rows = ws_raw.to_python()
        headers = raw_python_rows[0] if raw_python_rows else []
        rows = raw_python_rows[1:] if len(raw_python_rows) > 1 else []
    except Exception as e:
        log.warning(f"Calamine read failed: {e}. Falling back to openpyxl.")
        wb_in = openpyxl.load_workbook(str(input_file), data_only=True)
        ws_raw = wb_in['Raw'] if 'Raw' in wb_in.sheetnames else wb_in.active
        headers = [cell.value for cell in ws_raw[1]]
        rows = [list(row) for row in ws_raw.iter_rows(min_row=2, values_only=True)]
        wb_in.close()

    source_dc_idx = find_col(headers, ['source_dc', 'source dc', 'sourcedc', 'dc'])
    vms_status_idx = find_col(headers, ['vms status', 'vms_status', 'vmsstatus', 'status'])

    # Filter rows to allowed DCs using dc_config
    filtered_rows = [
        r for r in rows
        if len(r) > source_dc_idx and str(r[source_dc_idx] or '').strip().lower() in ALLOWED_DCS_SET_LOWER
    ]

    log.info(f"Filtered {len(filtered_rows)} rows matching allowed DCs")

    # Compute summary per Source_DC
    stats = defaultdict(lambda: {'done': 0, 'not_done': 0})

    for row in filtered_rows:
        if len(row) <= vms_status_idx:
            continue
        status = str(row[vms_status_idx] or '').strip().lower()
        dc_key = str(row[source_dc_idx] or '').strip().lower()
        if status == 'done':
            stats[dc_key]['done'] += 1
        else:
            stats[dc_key]['not_done'] += 1

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

    # ===== Summary sheet =====
    ws_sum = wb_out.active
    ws_sum.title = 'Summary'
    ws_sum.sheet_view.showGridLines = False

    # Fill cells white
    for r in range(1, len(summary_rows) + 10):
        for c in range(1, 8):
            ws_sum.cell(row=r, column=c).fill = white_fill

    # Row 1: Banner
    ws_sum.cell(row=1, column=1, value='VMS Adherence Summary')
    ws_sum.merge_cells(start_row=1, start_column=1, end_row=1, end_column=5)
    banner_cell = ws_sum.cell(row=1, column=1)
    banner_cell.fill = banner_fill
    banner_cell.font = white_font
    banner_cell.alignment = center
    ws_sum.row_dimensions[1].height = 22

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

    wb_out.save(str(output_file))
    log.info(f"Successfully generated VMS Adherence Report: {output_file}")


def main():
    script_dir = Path(__file__).resolve().parent
    input_file = script_dir.parent.parent / "vms adherence" / "VMS_Adherence_Report_30-Jul-2026.xlsx"
    output_file = script_dir.parent.parent / "vms adherence" / "output.xlsx"

    if len(sys.argv) > 1:
        input_file = Path(sys.argv[1])
    if len(sys.argv) > 2:
        output_file = Path(sys.argv[2])

    generate_vms_adherence_report(input_file, output_file)


if __name__ == "__main__":
    main()
