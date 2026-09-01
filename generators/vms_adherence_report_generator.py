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
    import gc
    log.info(f"Loading input workbook for VMS Adherence Report (streaming mode): {input_file}")

    wb_out = openpyxl.Workbook()
    wb_out.remove(wb_out.active)
    ws_raw_out = wb_out.create_sheet('Raw')

    headers = []
    source_dc_idx = 0
    vms_status_idx = 1
    stats = defaultdict(lambda: {'done': 0, 'not_done': 0})
    total_streamed = 0

    try:
        from python_calamine import CalamineWorkbook
        calamine_wb = CalamineWorkbook.from_path(str(input_file))
        ws_raw = calamine_wb.get_sheet_by_name('Raw') if 'Raw' in calamine_wb.sheet_names else calamine_wb.get_sheet_by_index(0)
        row_iter = iter(ws_raw.iter_rows())
        headers_raw = next(row_iter, None)
        if headers_raw:
            headers = list(headers_raw)
            source_dc_idx = find_col(headers, ['source_dc', 'source dc', 'sourcedc', 'dc'])
            vms_status_idx = find_col(headers, ['vms status', 'vms_status', 'vmsstatus', 'status'])
            ws_raw_out.append(headers)

            for row in row_iter:
                if len(row) > source_dc_idx and str(row[source_dc_idx] or '').strip().lower() in ALLOWED_DCS_SET_LOWER:
                    ws_raw_out.append(list(row))
                    total_streamed += 1
                    status = str(row[vms_status_idx] or '').strip().lower() if len(row) > vms_status_idx else ''
                    dc_key = str(row[source_dc_idx] or '').strip().lower()
                    if status == 'done':
                        stats[dc_key]['done'] += 1
                    else:
                        stats[dc_key]['not_done'] += 1
    except Exception as e:
        log.warning(f"Calamine stream failed: {e}. Falling back to openpyxl.")
        wb_in = openpyxl.load_workbook(str(input_file), data_only=True, read_only=True)
        ws_raw = wb_in['Raw'] if 'Raw' in wb_in.sheetnames else wb_in.active
        row_iter = ws_raw.iter_rows(values_only=True)
        headers_raw = next(row_iter, None)
        if headers_raw:
            headers = list(headers_raw)
            source_dc_idx = find_col(headers, ['source_dc', 'source dc', 'sourcedc', 'dc'])
            vms_status_idx = find_col(headers, ['vms status', 'vms_status', 'vmsstatus', 'status'])
            ws_raw_out.append(headers)

            for row in row_iter:
                if len(row) > source_dc_idx and str(row[source_dc_idx] or '').strip().lower() in ALLOWED_DCS_SET_LOWER:
                    ws_raw_out.append(list(row))
                    total_streamed += 1
                    status = str(row[vms_status_idx] or '').strip().lower() if len(row) > vms_status_idx else ''
                    dc_key = str(row[source_dc_idx] or '').strip().lower()
                    if status == 'done':
                        stats[dc_key]['done'] += 1
                    else:
                        stats[dc_key]['not_done'] += 1
        wb_in.close()

    log.info(f"Streamed {total_streamed} rows matching allowed DCs into Raw tab")

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

    banner_font = Font(name='Calibri', size=13, bold=True, color='FFFFFFFF')
    header_font = Font(name='Calibri', size=10, bold=True, color='FFFFFFFF')
    data_font = Font(name='Calibri', size=10, color='FF1e1b4b')
    total_font = Font(name='Calibri', size=10, bold=True, color='FF1e1b4b')
    green_font = Font(name='Calibri', size=10, bold=True, color='FF15803d')
    yellow_font = Font(name='Calibri', size=10, bold=True, color='FFa16207')
    red_font = Font(name='Calibri', size=10, bold=True, color='FFb91c1c')
    raw_header_font = Font(name='Calibri', size=10, bold=True, color='FFFFFFFF')

    center = Alignment(horizontal='center', vertical='center')
    align_left = Alignment(horizontal='left', vertical='center')
    align_right = Alignment(horizontal='right', vertical='center')

    # ===== Summary sheet =====
    ws_sum = wb_out.create_sheet('Summary', 0)
    ws_sum.sheet_view.showGridLines = True

    ws_sum.merge_cells('A1:E1')
    b_cell = ws_sum['A1']
    b_cell.value = 'VMS ADHERENCE SUMMARY'
    b_cell.font = banner_font
    b_cell.fill = banner_fill
    b_cell.alignment = center
    ws_sum.row_dimensions[1].height = 30

    for col in range(1, 6):
        cell = ws_sum.cell(row=1, column=col)
        cell.fill = banner_fill
        cell.border = purple_border

    ws_sum.row_dimensions[2].height = 8

    sum_headers = ['Source DC', 'Grand Total', 'Done', 'Not Done', 'Done %']
    for i, h in enumerate(sum_headers, 1):
        cell = ws_sum.cell(row=3, column=i, value=h)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = center
        cell.border = purple_border
    ws_sum.row_dimensions[3].height = 24

    for idx, srow in enumerate(summary_rows):
        row_num = idx + 4
        fill = white_fill if idx % 2 == 0 else alt_fill

        for c_idx, val in enumerate(srow, 1):
            cell = ws_sum.cell(row=row_num, column=c_idx, value=val)
            cell.font = data_font
            cell.fill = fill
            cell.border = data_border
            if c_idx == 1:
                cell.alignment = align_left
            elif c_idx == 5:
                cell.number_format = '0.0%'
                cell.alignment = center
            else:
                cell.number_format = '#,##0'
                cell.alignment = align_right

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

    tot_row = len(summary_rows) + 4
    tot_vals = ['Total', grand_total, grand_done, grand_not_done, grand_done_pct]
    for c_idx, val in enumerate(tot_vals, 1):
        cell = ws_sum.cell(row=tot_row, column=c_idx, value=val)
        cell.font = total_font
        cell.fill = alt_fill
        cell.border = purple_border
        if c_idx == 1:
            cell.alignment = align_left
        elif c_idx == 5:
            cell.number_format = '0.0%'
            cell.alignment = center
        else:
            cell.number_format = '#,##0'
            cell.alignment = align_right
    ws_sum.row_dimensions[tot_row].height = 22

    for col in range(1, 6):
        ws_sum.column_dimensions[get_column_letter(col)].width = 16

    # Style Raw header
    for i in range(1, len(headers) + 1):
        cell = ws_raw_out.cell(row=1, column=i)
        cell.fill = raw_header_fill
        cell.font = raw_header_font
        cell.alignment = center

    output_path = Path(output_file)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    wb_out.save(str(output_path))
    try:
        wb_out.close()
    except Exception:
        pass

    del wb_out
    gc.collect()
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
