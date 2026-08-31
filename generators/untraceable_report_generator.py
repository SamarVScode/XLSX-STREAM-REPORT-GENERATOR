#!/usr/bin/env python3
"""
Untraceable Stream Report Generator Module for ei_stream_server
==============================================================
Streams rows from Untraceable spreadsheet, pivots shipment count by Source DC & Age Bucket in-flight,
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

log = logging.getLogger("ei_stream_server.untraceable_report")

STANDARD_AGE_BUCKETS = [
    '0-2 Days', '3-5 Days', '6-10 Days', '11-20 Days', '21-30 Days', '>30 Days'
]

HIGHLIGHT_AGING_BUCKETS = {'6-10 Days', '11-20 Days', '21-30 Days', '>30 Days', '7-15 Days', '15-30 Days', '30-45 Days', '45+ Days'}

def generate_untraceable_report(input_file: Path, output_file: Path):
    path = Path(input_file)
    output_path = Path(output_file)
    log.info(f"Stream generating Untraceable Report: {path.name}")

    sheet_names = get_sheet_names(path)
    target_sheet = 'Raw' if 'Raw' in sheet_names else sheet_names[0]

    rows_iter = stream_sheet_rows(path, sheet_name=target_sheet)
    try:
        raw_headers = next(rows_iter)
    except StopIteration:
        raise ValueError(f"Sheet '{target_sheet}' is empty.")

    headers = [str(h).strip() if h is not None else f"Unnamed_{i}" for i, h in enumerate(raw_headers)]
    headers_lower = [h.lower() for h in headers]

    # Find columns
    sdc_idx = -1
    for cand in ['source dc', 'sourcedc', 'source_dc', 'dc']:
        if cand in headers_lower:
            sdc_idx = headers_lower.index(cand)
            break
    if sdc_idx == -1:
        sdc_idx = 0

    bucket_idx = -1
    for cand in ['age bucket', 'age_bucket', 'aging bucket']:
        if cand in headers_lower:
            bucket_idx = headers_lower.index(cand)
            break

    amt_idx = -1
    for cand in ['amount', 'shipmentamount']:
        if cand in headers_lower:
            amt_idx = headers_lower.index(cand)
            break

    filtered_rows = []
    dc_bucket_counts = defaultdict(lambda: defaultdict(int))
    dc_amounts = defaultdict(float)
    seen_buckets = set()

    for row in rows_iter:
        if not row or len(row) <= sdc_idx:
            continue
        sdc_raw = str(row[sdc_idx] or '').strip()
        if sdc_raw.lower() in ALLOWED_DCS_SET_LOWER:
            filtered_rows.append(row)
            sdc = sdc_raw.upper()
            
            b = str(row[bucket_idx] or '0-2 Days').strip() if bucket_idx != -1 and len(row) > bucket_idx else '0-2 Days'
            seen_buckets.add(b)
            dc_bucket_counts[sdc][b] += 1

            if amt_idx != -1 and len(row) > amt_idx and row[amt_idx] is not None:
                try:
                    dc_amounts[sdc] += float(row[amt_idx])
                except (ValueError, TypeError):
                    pass

    log.info(f"Streamed and filtered {len(filtered_rows)} matching Untraceable rows.")

    present_buckets = [b for b in STANDARD_AGE_BUCKETS if b in seen_buckets]
    other_buckets = [b for b in seen_buckets if b not in STANDARD_AGE_BUCKETS]
    all_buckets = present_buckets + sorted(other_buckets)
    if not all_buckets:
        all_buckets = STANDARD_AGE_BUCKETS

    active_dcs = [dc for dc in dc_bucket_counts if sum(dc_bucket_counts[dc].values()) > 0]
    active_dcs.sort()

    wb = openpyxl.Workbook()

    header_fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
    alt_row_fill = PatternFill(start_color="F9FAFB", end_color="F9FAFB", fill_type="solid")
    aging_highlight_fill = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
    font_aging_highlight = Font(name="Calibri", size=11, bold=True, color="9C0006")
    font_header = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    font_regular = Font(name="Calibri", size=11)

    thin_border_side = Side(border_style="thin", color="D9D9D9")
    border_cell = Border(left=thin_border_side, right=thin_border_side, top=thin_border_side, bottom=thin_border_side)

    # ----------------------------------------------------
    # Sheet 1: Summary
    # ----------------------------------------------------
    ws_sum = wb.active
    ws_sum.title = "Summary"
    ws_sum.sheet_view.showGridLines = False

    num_buckets = len(all_buckets)
    age_start_col = 2
    age_end_col = 1 + num_buckets
    amt_col_idx = 2 + num_buckets

    ws_sum.row_dimensions[1].height = 22
    ws_sum.row_dimensions[2].height = 22

    ws_sum.merge_cells(start_row=1, start_column=1, end_row=2, end_column=1)
    ws_sum.cell(row=1, column=1, value="Source DC")
    for r in range(1, 3):
        c = ws_sum.cell(row=r, column=1)
        c.font = font_header
        c.fill = header_fill
        c.alignment = Alignment(horizontal="center", vertical="center")
        c.border = border_cell

    ws_sum.merge_cells(start_row=1, start_column=age_start_col, end_row=1, end_column=age_end_col)
    ws_sum.cell(row=1, column=age_start_col, value="Age Bucket")
    for col_idx in range(age_start_col, age_end_col + 1):
        c = ws_sum.cell(row=1, column=col_idx)
        c.font = font_header
        c.fill = header_fill
        c.alignment = Alignment(horizontal="center", vertical="center")
        c.border = border_cell

    for idx, b_text in enumerate(all_buckets):
        col_idx = age_start_col + idx
        c = ws_sum.cell(row=2, column=col_idx, value=b_text)
        c.font = font_header
        c.fill = header_fill
        c.alignment = Alignment(horizontal="center", vertical="center")
        c.border = border_cell

    ws_sum.merge_cells(start_row=1, start_column=amt_col_idx, end_row=2, end_column=amt_col_idx)
    ws_sum.cell(row=1, column=amt_col_idx, value="Amount")
    for r in range(1, 3):
        c = ws_sum.cell(row=r, column=amt_col_idx)
        c.font = font_header
        c.fill = header_fill
        c.alignment = Alignment(horizontal="center", vertical="center")
        c.border = border_cell

    row_idx = 3
    for dc_name in active_dcs:
        c1 = ws_sum.cell(row=row_idx, column=1, value=str(dc_name))
        c1.font = font_regular
        c1.border = border_cell

        for col_idx, b_name in enumerate(all_buckets, start=2):
            cnt = dc_bucket_counts[dc_name].get(b_name, 0)
            cell_val = cnt if cnt > 0 else ""
            c = ws_sum.cell(row=row_idx, column=col_idx, value=cell_val)
            c.alignment = Alignment(horizontal="right", vertical="center")
            c.border = border_cell
            if cell_val != "":
                c.number_format = "#,##0"

            if b_name in HIGHLIGHT_AGING_BUCKETS and cnt > 0:
                c.fill = aging_highlight_fill
                c.font = font_aging_highlight
            else:
                c.font = font_regular

        amt_val = dc_amounts.get(dc_name, 0.0)
        c_amt = ws_sum.cell(row=row_idx, column=amt_col_idx, value=amt_val if amt_val > 0 else "")
        c_amt.alignment = Alignment(horizontal="right", vertical="center")
        c_amt.border = border_cell
        c_amt.font = font_regular
        if amt_val > 0:
            c_amt.number_format = "#,##0"

        row_idx += 1

    for col in ws_sum.columns:
        max_len = 0
        col_letter = get_column_letter(col[0].column)
        for cell in col:
            val_str = str(cell.value or '')
            if cell.number_format and '#' in cell.number_format and isinstance(cell.value, (int, float)):
                val_str = f"{cell.value:,.0f}"
            max_len = max(max_len, len(val_str))
        ws_sum.column_dimensions[col_letter].width = max(max_len + 4, 12)

    # ----------------------------------------------------
    # Sheet 2: Raw
    # ----------------------------------------------------
    ws_raw = wb.create_sheet(title="Raw")
    ws_raw.sheet_view.showGridLines = True

    for col_idx, h_text in enumerate(headers, start=1):
        cell = ws_raw.cell(row=1, column=col_idx, value=h_text)
        cell.font = font_header
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = border_cell
    ws_raw.row_dimensions[1].height = 25

    pincode_idx = headers_lower.index('customerpincode') if 'customerpincode' in headers_lower else -1

    for r_idx, row_values in enumerate(filtered_rows, start=2):
        use_alt = (r_idx % 2 == 0)
        for col_idx, val in enumerate(row_values, start=1):
            cell = ws_raw.cell(row=r_idx, column=col_idx)
            
            if val is None or val == '':
                cell.value = ""
            elif isinstance(val, (int, float)):
                cell.value = val
                if col_idx - 1 == amt_idx:
                    cell.number_format = "#,##0"
                    cell.alignment = Alignment(horizontal="right")
                elif col_idx - 1 == pincode_idx:
                    cell.number_format = "0"
                    cell.alignment = Alignment(horizontal="center")
                else:
                    cell.alignment = Alignment(horizontal="right")
            else:
                cell.value = str(val)
                cell.alignment = Alignment(horizontal="left")

            cell.font = font_regular
            cell.border = border_cell
            if use_alt:
                cell.fill = alt_row_fill

    for col in ws_raw.columns:
        max_len = 0
        col_letter = get_column_letter(col[0].column)
        for cell in col:
            val_str = str(cell.value or '')
            if cell.number_format and '#' in cell.number_format and isinstance(cell.value, (int, float)):
                val_str = f"{cell.value:,.0f}"
            max_len = max(max_len, len(val_str))
        ws_raw.column_dimensions[col_letter].width = min(max(max_len + 3, 12), 40)

    output_path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(output_path)
    log.info(f"Successfully generated Untraceable Report: {output_path}")
    return str(output_path)
