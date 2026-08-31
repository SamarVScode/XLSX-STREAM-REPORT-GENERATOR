#!/usr/bin/env python3
"""
Reverse Pendency Stream Generator Module for ei_stream_server
=============================================================
Streams rows from input spreadsheet, computes Age_Bucket in-flight, and generates output workbook.
"""

import logging
from pathlib import Path
from collections import defaultdict
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter

from config.dc_config import ALLOWED_SOURCE_DCS, ALLOWED_DCS_SET
from core.stream_engine import stream_sheet_rows, get_sheet_names

AGING_CATEGORIES = ['0-2 Days', '3-5 Days', '6-10 Days', '>10 Days']

log = logging.getLogger("ei_stream_server.reverse_pendency")

def compute_age_bucket(val) -> str:
    if val is None or str(val).strip() == '':
        return '0-2 Days'
    try:
        aging = float(val)
        if aging <= 2:
            return '0-2 Days'
        elif aging <= 5:
            return '3-5 Days'
        elif aging <= 10:
            return '6-10 Days'
        else:
            return '>10 Days'
    except (ValueError, TypeError):
        return '0-2 Days'

def build_summary_sheet(out_wb, filtered_rows, src_dc_idx, age_bucket_idx):
    ws = out_wb.create_sheet(title="Summary", index=0)
    ws.sheet_view.showGridLines = False

    title_fill = PatternFill("solid", fgColor="1E1B4B")
    title_font = Font(name="Calibri", size=12, bold=True, color="FFFFFF")
    header_fill = PatternFill("solid", fgColor="312E81")
    header_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    total_fill = PatternFill("solid", fgColor="E0E7FF")
    total_font = Font(name="Calibri", size=11, bold=True, color="1E1B4B")
    data_font = Font(name="Calibri", size=11, color="1F2937")
    red_fill = PatternFill("solid", fgColor="FECACA")
    red_font = Font(name="Calibri", size=11, bold=True, color="991B1B")
    thin_side = Side(style="thin", color="CBD5E1")
    border = Border(left=thin_side, right=thin_side, top=thin_side, bottom=thin_side)
    center = Alignment(horizontal="center", vertical="center")
    left = Alignment(horizontal="left", vertical="center")

    pivot = defaultdict(lambda: defaultdict(int))
    for row in filtered_rows[1:]:
        dc = str(row[src_dc_idx]).strip().upper() if len(row) > src_dc_idx and row[src_dc_idx] else ''
        bucket = str(row[age_bucket_idx]).strip() if len(row) > age_bucket_idx and row[age_bucket_idx] else '0-2 Days'
        pivot[dc][bucket] += 1

    headers = ["Source DC"] + AGING_CATEGORIES + ["Total Pendency"]
    data = []
    tot_buckets = defaultdict(int)

    for dc in ALLOWED_SOURCE_DCS:
        row_vals = [dc]
        row_tot = 0
        for cat in AGING_CATEGORIES:
            cnt = pivot[dc][cat]
            row_vals.append(cnt)
            row_tot += cnt
            tot_buckets[cat] += cnt
        row_vals.append(row_tot)
        data.append(row_vals)

    totals = ["Total"] + [tot_buckets[cat] for cat in AGING_CATEGORIES] + [sum(tot_buckets.values())]
    data.append(totals)

    end_col = len(headers)
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=end_col)
    title_cell = ws.cell(row=1, column=1, value="Aging wise report")
    title_cell.font = title_font
    title_cell.alignment = center
    for c in range(1, end_col + 1):
        cell = ws.cell(row=1, column=c)
        cell.fill = title_fill
        cell.border = border

    for i, h in enumerate(headers):
        cell = ws.cell(row=2, column=i + 1, value=h)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = center
        cell.border = border

    for r_idx, row_vals in enumerate(data):
        row_num = r_idx + 3
        is_total = row_vals[0] == "Total"
        for c_idx, val in enumerate(row_vals):
            cell = ws.cell(row=row_num, column=c_idx + 1, value=val)
            cell.border = border
            cell.alignment = left if c_idx == 0 else center

            if isinstance(val, (int, float)) and val > 0 and headers[c_idx] in ['3-5 Days', '6-10 Days', '>10 Days']:
                cell.fill = red_fill
                cell.font = red_font
            elif is_total:
                cell.fill = total_fill
                cell.font = total_font
            else:
                cell.font = data_font

    for col in range(1, end_col + 1):
        letter = get_column_letter(col)
        max_len = max(len(str(ws.cell(r, col).value or '')) for r in range(1, len(data) + 3))
        ws.column_dimensions[letter].width = max(max_len + 3, 14)

def build_p0_sheet(out_wb, filtered_rows, header):
    ws = out_wb.create_sheet(title="Critical P0")
    ws.sheet_view.showGridLines = False

    title_fill = PatternFill("solid", fgColor="1E1B4B")
    title_font = Font(name="Calibri", size=12, bold=True, color="FFFFFF")
    header_fill = PatternFill("solid", fgColor="312E81")
    header_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    data_font = Font(name="Calibri", size=11, color="1F2937")
    thin_side = Side(style="thin", color="CBD5E1")
    border = Border(left=thin_side, right=thin_side, top=thin_side, bottom=thin_side)
    center = Alignment(horizontal="center", vertical="center")
    left = Alignment(horizontal="left", vertical="center")

    tn_idx = header.index('tracking_number') if 'tracking_number' in header else 0
    src_dc_idx = header.index('Source DC') if 'Source DC' in header else 1
    aging_idx = header.index('Aging') if 'Aging' in header else 3
    age_bucket_idx = header.index('Age_Bucket') if 'Age_Bucket' in header else 4
    attempt_idx = header.index('Attempt_Status') if 'Attempt_Status' in header else 5

    p0_rows = []
    for row in filtered_rows[1:]:
        try:
            aging = float(row[aging_idx]) if len(row) > aging_idx and row[aging_idx] else 0
        except (ValueError, TypeError):
            aging = 0
        if aging >= 2:
            p0_rows.append(row)

    out_headers = ['tracking_number', 'Source DC', 'Aging', 'Age_Bucket', 'Attempt_Status']

    end_col = len(out_headers)
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=end_col)
    title_cell = ws.cell(row=1, column=1, value="P0 reverse pendency (Aging ≥ 2)")
    title_cell.font = title_font
    title_cell.alignment = center
    for c in range(1, end_col + 1):
        cell = ws.cell(row=1, column=c)
        cell.fill = title_fill
        cell.border = border

    for i, h in enumerate(out_headers):
        cell = ws.cell(row=2, column=i + 1, value=h)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = center
        cell.border = border

    for r_idx, row in enumerate(p0_rows):
        row_num = r_idx + 3
        vals = [
            row[tn_idx] if len(row) > tn_idx else '',
            str(row[src_dc_idx]).strip().upper() if len(row) > src_dc_idx and row[src_dc_idx] else '',
            row[aging_idx] if len(row) > aging_idx else '',
            str(row[age_bucket_idx]).strip() if len(row) > age_bucket_idx and row[age_bucket_idx] else '',
            str(row[attempt_idx]).strip() if len(row) > attempt_idx and row[attempt_idx] else ''
        ]
        for c_idx, val in enumerate(vals):
            cell = ws.cell(row=row_num, column=c_idx + 1, value=val)
            cell.border = border
            cell.font = data_font
            cell.alignment = left if c_idx in [0, 4] else center

    col_widths = [20, 14, 10, 14, 22]
    for i, w in enumerate(col_widths):
        ws.column_dimensions[get_column_letter(i + 1)].width = w

    return len(p0_rows)

def generate_reverse_pendency_report(input_file: Path, output_file: Path):
    path = Path(input_file)
    log.info(f"Stream generating Reverse Pendency Report: {path}")
    sheet_names = get_sheet_names(path)

    target_sheet = 'Raw'
    if 'Raw' not in sheet_names:
        sheet_map = {name.lower(): name for name in sheet_names}
        if 'raw' in sheet_map:
            target_sheet = sheet_map['raw']
        elif len(sheet_names) > 0:
            target_sheet = sheet_names[0]

    rows_iter = stream_sheet_rows(path, sheet_name=target_sheet)
    try:
        raw_header = next(rows_iter)
    except StopIteration:
        raise ValueError(f"Sheet '{target_sheet}' is empty.")

    header = [str(h).strip() if h is not None else '' for h in raw_header]

    src_dc_idx = header.index('Source DC') if 'Source DC' in header else 1
    region_idx = header.index('Region') if 'Region' in header else 2
    aging_idx = header.index('Aging') if 'Aging' in header else 3
    
    if 'Age_Bucket' not in header:
        header.insert(aging_idx + 1, 'Age_Bucket')
        age_bucket_idx = aging_idx + 1
    else:
        age_bucket_idx = header.index('Age_Bucket')

    filtered = [header]
    for row in rows_iter:
        if not row:
            continue
        region = str(row[region_idx]).strip() if len(row) > region_idx and row[region_idx] else ''
        src_dc = str(row[src_dc_idx]).strip().upper() if len(row) > src_dc_idx and row[src_dc_idx] else ''
        
        # If Region is present, check North; otherwise match allowed DCs
        region_matches = (region.lower() == 'north') if region else True
        if region_matches and src_dc in ALLOWED_SOURCE_DCS:
            row_list = list(row)
            if len(row_list) <= age_bucket_idx:
                row_list.extend([None] * (age_bucket_idx - len(row_list) + 1))
            aging_val = row_list[aging_idx] if len(row_list) > aging_idx else None
            bucket = compute_age_bucket(aging_val)
            row_list[age_bucket_idx] = bucket
            filtered.append(row_list)

    log.info(f"Streamed and filtered {len(filtered) - 1} matching rows.")

    out_wb = Workbook()
    ws = out_wb.active
    ws.title = "Raw"
    for row in filtered:
        ws.append(row)

    build_summary_sheet(out_wb, filtered, src_dc_idx, age_bucket_idx)
    p0_count = build_p0_sheet(out_wb, filtered, header)

    out_wb.save(str(output_file))
    log.info(f"Saved Reverse Pendency Report: {output_file.name} ({len(filtered)} rows, {p0_count} P0 rows)")
    return str(output_file)
