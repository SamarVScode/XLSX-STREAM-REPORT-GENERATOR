#!/usr/bin/env python3
"""
EOB Stream Report Generator Module for ei_stream_server
========================================================
Streams rows from input spreadsheet, computes Ageing Bucket and Status metrics per DC in-flight,
and generates formatted output workbook.
"""

import logging
from pathlib import Path
from collections import defaultdict
import pandas as pd
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

from config.dc_config import ALLOWED_SOURCE_DCS
from core.stream_engine import stream_sheet_rows, get_sheet_names

log = logging.getLogger("ei_stream_server.eob_generator")

TARGET_SOURCE_DCS = [dc.upper() for dc in ALLOWED_SOURCE_DCS if str(dc).upper() != 'ALL']

STATUS_SHORTFORMS = {
    'Out_For_Delivery': 'OFD',
    'Undelivered_HeavyLoad': 'Heavy Load',
    'Undelivered_No_Response': 'No Response',
    'Undelivered_NonServiceablePincode': 'Non-Serv Pincode',
    'Undelivered_Not_Attended': 'Not Attended',
    'Undelivered_Request_For_Reschedule': 'Reschedule',
    'Undelivered_Request_For_Reschedule_Customer_Triggered': 'Cust Reschedule',
    'Untraceable': 'Untraceable',
    'Undelivered_SameStateMisroute': 'Same State Misroute',
    'Undelivered_OtherStateMisroute': 'Other State Misroute',
    'Undelivered_Order_Rejected_By_Customer': 'Rejected',
    'Undelivered_Order_Rejected_By_Customer_Customer_Triggered': 'Cust Rejected',
    'Undelivered_Incomplete_Address': 'Incomplete Addr',
    'Undelivered_Attempted': 'Attempted',
    'Undelivered_UntraceableFromHub': 'Untraced Hub',
    'Untraceable_BRSNR': 'Untraceable BRSNR',
}

def get_short_status(status_str):
    if not status_str or pd.isna(status_str):
        return 'Unknown'
    s = str(status_str).strip()
    return STATUS_SHORTFORMS.get(s, s)

def generate_eob_report(input_file: Path, output_file: Path):
    path = Path(input_file)
    output_path = Path(output_file)
    log.info(f"Stream generating EOB Report: {path.name}")

    sheet_names = get_sheet_names(path)
    sheet_map = {name.lower(): name for name in sheet_names}
    target_sheet = None
    for candidate in ['raw', 'raw_data', 'raw_data_north', 'praw data']:
        if candidate in sheet_map:
            target_sheet = sheet_map[candidate]
            break
    if not target_sheet:
        target_sheet = sheet_names[0] if sheet_names else 'Raw'

    rows_iter = stream_sheet_rows(path, sheet_name=target_sheet)
    try:
        raw_headers = next(rows_iter)
    except StopIteration:
        raise ValueError(f"Sheet '{target_sheet}' is empty.")

    headers = [str(h).strip() if h is not None else f"Unnamed_{i}" for i, h in enumerate(raw_headers)]
    headers_lower = [h.lower() for h in headers]

    # Find columns
    sdc_idx = -1
    for cand in ['source dc', 'source_dc', 'dc']:
        if cand in headers_lower:
            sdc_idx = headers_lower.index(cand)
            break
    if sdc_idx == -1:
        sdc_idx = 0

    status_idx = -1
    for cand in ['latest status', 'latest_status', 'status_status', 'status']:
        if cand in headers_lower:
            status_idx = headers_lower.index(cand)
            break

    ageing_idx = -1
    for cand in ['ageing bucket', 'ageing_bucket', 'aging bucket', 'aging_bucket']:
        if cand in headers_lower:
            ageing_idx = headers_lower.index(cand)
            break

    filtered_rows = []
    t1_counts = defaultdict(lambda: defaultdict(int))
    t2_counts = defaultdict(lambda: defaultdict(int))
    seen_buckets = set()
    seen_statuses = set()

    for row in rows_iter:
        if not row or len(row) <= sdc_idx:
            continue
        sdc = str(row[sdc_idx] or '').strip().upper()
        if sdc in TARGET_SOURCE_DCS:
            filtered_rows.append(row)
            
            b = str(row[ageing_idx] or 'Unknown').strip() if ageing_idx != -1 and len(row) > ageing_idx else 'Unknown'
            s_raw = str(row[status_idx] or 'Unknown').strip() if status_idx != -1 and len(row) > status_idx else 'Unknown'
            s = get_short_status(s_raw)

            t1_counts[sdc][b] += 1
            t2_counts[sdc][s] += 1
            seen_buckets.add(b)
            seen_statuses.add(s)

    log.info(f"Streamed and filtered {len(filtered_rows)} matching EOB rows.")

    ageing_order = ['1-2 days', '3-5 days', '6-10 days', '11-15 days', '16-20 days', '>20 days']
    sorted_buckets = [b for b in ageing_order if b in seen_buckets]
    for b in sorted(seen_buckets):
        if b not in sorted_buckets:
            sorted_buckets.append(b)

    sorted_statuses = sorted(list(seen_statuses))

    # --- BUILD WORKBOOK ---
    wb = openpyxl.Workbook()
    ws_summary = wb.active
    ws_summary.title = "Summary"
    ws_summary.sheet_view.showGridLines = False

    font_header = Font(name='Calibri', size=10, bold=True, color='FFFFFF')
    font_data = Font(name='Calibri', size=10, color='1E293B')
    font_total = Font(name='Calibri', size=10, bold=True, color='0F172A')
    font_priority = Font(name='Calibri', size=10, bold=True, color='991B1B')
    fill_priority = PatternFill(start_color='FEE2E2', end_color='FEE2E2', fill_type='solid')
    fill_header = PatternFill(start_color='1E293B', end_color='1E293B', fill_type='solid')
    fill_total = PatternFill(start_color='F1F5F9', end_color='F1F5F9', fill_type='solid')
    fill_zebra = PatternFill(start_color='F8FAFC', end_color='F8FAFC', fill_type='solid')

    border_thin = Side(style='thin', color='CBD5E1')
    border_double = Side(style='double', color='475569')
    border_cell = Border(left=border_thin, right=border_thin, top=border_thin, bottom=border_thin)
    border_header = Border(left=border_thin, right=border_thin, top=border_thin, bottom=border_thin)
    border_total = Border(left=border_thin, right=border_thin, top=border_thin, bottom=border_double)

    align_center = Alignment(horizontal='center', vertical='center')
    align_right = Alignment(horizontal='right', vertical='center')

    start_r = 1
    start_c1 = 1

    # Write Table 1
    headers_t1 = ['Source DC'] + sorted_buckets + ['Grand Total']
    for c_idx, h in enumerate(headers_t1, start=start_c1):
        cell = ws_summary.cell(row=start_r, column=c_idx, value=h)
        cell.font = font_header
        cell.fill = fill_header
        cell.alignment = align_center
        cell.border = border_header

    tot_buckets_t1 = defaultdict(int)
    for r_offset, dc in enumerate(TARGET_SOURCE_DCS, start=start_r + 1):
        c_cell = ws_summary.cell(row=r_offset, column=start_c1, value=dc)
        c_cell.font = font_data
        c_cell.alignment = align_center
        c_cell.border = border_cell

        row_tot = 0
        for c_idx, b in enumerate(sorted_buckets, start=start_c1 + 1):
            cnt = t1_counts[dc][b]
            row_tot += cnt
            tot_buckets_t1[b] += cnt

            cell = ws_summary.cell(row=r_offset, column=c_idx, value=cnt)
            cell.alignment = align_right
            cell.border = border_cell
            if cnt > 0:
                cell.fill = fill_priority
                cell.font = font_priority
            else:
                cell.font = font_data
                if r_offset % 2 == 0:
                    cell.fill = fill_zebra

        tot_cell = ws_summary.cell(row=r_offset, column=start_c1 + len(sorted_buckets) + 1, value=row_tot)
        tot_cell.font = font_data
        tot_cell.alignment = align_right
        tot_cell.border = border_cell

    # Total Row Table 1
    t1_tot_row_num = start_r + len(TARGET_SOURCE_DCS) + 1
    t1_tot_label = ws_summary.cell(row=t1_tot_row_num, column=start_c1, value='Grand Total')
    t1_tot_label.font = font_total
    t1_tot_label.fill = fill_total
    t1_tot_label.alignment = align_center
    t1_tot_label.border = border_total

    g_total_t1 = 0
    for c_idx, b in enumerate(sorted_buckets, start=start_c1 + 1):
        b_sum = tot_buckets_t1[b]
        g_total_t1 += b_sum
        cell = ws_summary.cell(row=t1_tot_row_num, column=c_idx, value=b_sum)
        cell.font = font_total
        cell.fill = fill_total
        cell.alignment = align_right
        cell.border = border_total

    g_tot_cell = ws_summary.cell(row=t1_tot_row_num, column=start_c1 + len(sorted_buckets) + 1, value=g_total_t1)
    g_tot_cell.font = font_total
    g_tot_cell.fill = fill_total
    g_tot_cell.alignment = align_right
    g_tot_cell.border = border_total

    # Write Table 2 (Side-by-Side)
    start_c2 = start_c1 + len(headers_t1) + 1
    headers_t2 = ['Source DC'] + sorted_statuses + ['Grand Total']

    for c_idx, h in enumerate(headers_t2, start=start_c2):
        cell = ws_summary.cell(row=start_r, column=c_idx, value=h)
        cell.font = font_header
        cell.fill = fill_header
        cell.alignment = align_center
        cell.border = border_header

    tot_stat_t2 = defaultdict(int)
    for r_offset, dc in enumerate(TARGET_SOURCE_DCS, start=start_r + 1):
        c_cell = ws_summary.cell(row=r_offset, column=start_c2, value=dc)
        c_cell.font = font_data
        c_cell.alignment = align_center
        c_cell.border = border_cell

        row_tot = 0
        for c_idx, s in enumerate(sorted_statuses, start=start_c2 + 1):
            cnt = t2_counts[dc][s]
            row_tot += cnt
            tot_stat_t2[s] += cnt

            cell = ws_summary.cell(row=r_offset, column=c_idx, value=cnt)
            cell.alignment = align_right
            cell.border = border_cell
            cell.font = font_data
            if r_offset % 2 == 0:
                cell.fill = fill_zebra

        tot_cell = ws_summary.cell(row=r_offset, column=start_c2 + len(sorted_statuses) + 1, value=row_tot)
        tot_cell.font = font_data
        tot_cell.alignment = align_right
        tot_cell.border = border_cell

    # Total Row Table 2
    t2_tot_row_num = start_r + len(TARGET_SOURCE_DCS) + 1
    t2_tot_label = ws_summary.cell(row=t2_tot_row_num, column=start_c2, value='Grand Total')
    t2_tot_label.font = font_total
    t2_tot_label.fill = fill_total
    t2_tot_label.alignment = align_center
    t2_tot_label.border = border_total

    g_total_t2 = 0
    for c_idx, s in enumerate(sorted_statuses, start=start_c2 + 1):
        s_sum = tot_stat_t2[s]
        g_total_t2 += s_sum
        cell = ws_summary.cell(row=t2_tot_row_num, column=c_idx, value=s_sum)
        cell.font = font_total
        cell.fill = fill_total
        cell.alignment = align_right
        cell.border = border_total

    g_tot_cell2 = ws_summary.cell(row=t2_tot_row_num, column=start_c2 + len(sorted_statuses) + 1, value=g_total_t2)
    g_tot_cell2.font = font_total
    g_tot_cell2.fill = fill_total
    g_tot_cell2.alignment = align_right
    g_tot_cell2.border = border_total

    # Column Widths
    gap_col_idx = start_c1 + len(headers_t1)
    gap_col_letter = get_column_letter(gap_col_idx)

    for col in ws_summary.columns:
        col_letter = get_column_letter(col[0].column)
        if col_letter == gap_col_letter:
            ws_summary.column_dimensions[col_letter].width = 4
            continue

        max_len = 0
        for cell in col:
            if cell.value is not None:
                max_len = max(max_len, len(str(cell.value)))
        ws_summary.column_dimensions[col_letter].width = max(max_len + 3, 11)

    # 2. Raw Sheet
    ws_raw = wb.create_sheet(title="Raw")
    ws_raw.sheet_view.showGridLines = True

    for c_idx, h in enumerate(headers, start=1):
        cell = ws_raw.cell(row=1, column=c_idx, value=h)
        cell.font = font_header
        cell.fill = fill_header
        cell.alignment = align_center

    for r_idx, row in enumerate(filtered_rows, start=2):
        for c_idx, val in enumerate(row, start=1):
            ws_raw.cell(row=r_idx, column=c_idx, value=val)

    output_path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(output_path)
    log.info(f"Successfully generated EOB Report: {output_path.name}")
    return str(output_path)
