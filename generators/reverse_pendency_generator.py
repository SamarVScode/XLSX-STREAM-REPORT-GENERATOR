#!/usr/bin/env python3
"""
Reverse Pendency Report Generator Module for ei_report_server
=============================================================
Reads 'Raw' sheet from input Excel file, filters rows where Region == 'North'
and Source DC is in allowed list, computes Age_Bucket, and generates output workbook:
  1. Summary Sheet (Aging wise report)
  2. P0 reverse pendency Sheet (Aging >= 2 tracking details)
  3. Raw Sheet (Full filtered dataset)
"""

import logging
from pathlib import Path
from collections import defaultdict
from python_calamine import CalamineWorkbook
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter

try:
    from config.dc_config import ALLOWED_SOURCE_DCS, ALLOWED_DCS_SET
except ImportError:
    from dc_config import ALLOWED_SOURCE_DCS, ALLOWED_DCS_SET
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
        dc = str(row[src_dc_idx]).strip().upper() if row[src_dc_idx] else ''
        bucket = str(row[age_bucket_idx]).strip() if row[age_bucket_idx] else '0-2 Days'
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

    col_map = {str(h).strip().lower(): idx for idx, h in enumerate(header) if h is not None}
    tn_idx = 0
    for c in ['tracking_number', 'tracking_no', 'tracking_id', 'tracking id', 'waybill', 'awb', 'shipment']:
        if c in col_map:
            tn_idx = col_map[c]
            break

    src_dc_idx = col_map.get('source dc', col_map.get('source_dc', col_map.get('dc', 0)))
    aging_idx = col_map.get('aging', col_map.get('age', 1))
    age_bucket_idx = col_map.get('age_bucket', col_map.get('age bucket', len(header) - 1))
    attempt_idx = col_map.get('attempt_status', col_map.get('status', col_map.get('attempt', 0)))

    p0_rows = []
    for row in filtered_rows[1:]:
        try:
            aging = float(row[aging_idx]) if row[aging_idx] else 0
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
            row[tn_idx],
            str(row[src_dc_idx]).strip().upper() if row[src_dc_idx] else '',
            row[aging_idx],
            str(row[age_bucket_idx]).strip() if row[age_bucket_idx] else '',
            str(row[attempt_idx]).strip() if row[attempt_idx] else ''
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
    log.info(f"Reading workbook for Reverse Pendency: {input_file}")
    wb = CalamineWorkbook.from_path(str(input_file))

    target_sheet = None
    sheet_map = {name.lower(): name for name in wb.sheet_names}
    for cand in ['raw', 'raw_data', 'data']:
        if cand in sheet_map:
            target_sheet = sheet_map[cand]
            break
    if not target_sheet:
        target_sheet = wb.sheet_names[0]

    sheet = wb.get_sheet_by_name(target_sheet)
    rows = list(sheet.iter_rows())
    if not rows:
        raise ValueError(f"Sheet '{target_sheet}' is empty.")
    header = rows[0]

    log.info(f"Raw sheet: {len(rows)-1} data rows, {len(header)} columns")

    col_map = {str(h).strip().lower(): idx for idx, h in enumerate(header) if h is not None}
    
    src_dc_idx = 0
    for c in ['source dc', 'source_dc', 'dc']:
        if c in col_map:
            src_dc_idx = col_map[c]
            break

    region_idx = col_map.get('region', None)
    
    aging_idx = 1
    for c in ['aging', 'aging bucket', 'age', 'age_bucket', 'ageing']:
        if c in col_map:
            aging_idx = col_map[c]
            break

    age_bucket_idx = col_map.get('age_bucket') or col_map.get('aging bucket') or col_map.get('age bucket')
    if age_bucket_idx is None:
        age_bucket_idx = len(header)
        header = list(header) + ['Age_Bucket']

    filtered = [header]
    for row in rows[1:]:
        region = str(row[region_idx]).strip() if (region_idx is not None and len(row) > region_idx and row[region_idx]) else 'North'
        src_dc = str(row[src_dc_idx]).strip().upper() if (len(row) > src_dc_idx and row[src_dc_idx]) else ''
        if (region == 'North' or region_idx is None) and src_dc in ALLOWED_SOURCE_DCS:
            row_list = list(row)
            while len(row_list) <= age_bucket_idx:
                row_list.append('')
            aging_val = row_list[aging_idx] if len(row_list) > aging_idx else 0
            bucket = compute_age_bucket(aging_val)
            row_list[age_bucket_idx] = bucket
            filtered.append(row_list)

    log.info(f"Filtered: {len(filtered)-1} rows (Region=North, Source DC in target list)")

    out_wb = Workbook()
    ws = out_wb.active
    ws.title = "Raw"
    for row in filtered:
        ws.append(row)

    build_summary_sheet(out_wb, filtered, src_dc_idx, age_bucket_idx)
    p0_count = build_p0_sheet(out_wb, filtered, header)

    out_wb.save(str(output_file))
    try:
        out_wb.close()
    except Exception:
        pass

    del out_wb
    import gc
    gc.collect()
    log.info(f"Saved Reverse Pendency Report: {output_file.name} ({len(filtered)} total rows incl header, {p0_count} P0 rows)")
