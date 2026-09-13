#!/usr/bin/env python3
"""
Weekly SCM TAT & HCQ Report Generator Module for ei_stream_server
================================================================
Generates a structured 6-tab workbook for SCM TAT & HCQ:
  1. 'summary'              : Filtered & aligned DC scorecard with Green/Yellow/Red coloring
  2. 'raw data'             : All rows for configured DCs from the 'Data' tab (Streamed to Disk XML)
  3. 'todays tasks'         : Open/pending tasks (status != 'Closed') (Streamed to Disk XML)
  4. "today's task summary" : Executive KPI cards (Open, Forward, Reverse), DC x Aging matrix,
                              and side-by-side L4 & L5 Root Cause breakdowns.
  5. 'hub l5 summary'       : Hub-wise matrix with distinct L5 Root Reasons as column headers
  6. 'hub l4 l5 breakdown'  : Detailed drilldown per hub (L4, L5, Open Tasks, Share of Hub %, Total Share %)

Uses Single-Pass Zero-Memory Streaming Engine (core.stream_engine):
- O(1) Memory Footprint (< 35MB RAM)
- Direct XML disk streaming for massive (10MB - 40MB+) datasets
"""

import sys
import logging
from pathlib import Path
from collections import defaultdict, Counter
import re
import datetime
import html
import zipfile
from typing import Dict, Any, List, Tuple

import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# Ensure server root is in sys.path
SERVER_ROOT = Path(__file__).resolve().parent.parent
if str(SERVER_ROOT) not in sys.path:
    sys.path.insert(0, str(SERVER_ROOT))

try:
    from config.dc_config import ALLOWED_SOURCE_DCS, ALLOWED_DCS_SET_LOWER, ALLOWED_DCS_SET, normalize_dc_code, is_allowed_dc
except ImportError:
    ALLOWED_SOURCE_DCS = [
        'GZB', 'NDA', 'GND', 'WDL', 'MEE',
        'AGR', 'MTH', 'SPR', 'MZN', 'HPA',
        'FZD', 'LKO', 'BRL', 'MOR', 'GKP',
        'DRD', 'HDN', 'HRD', 'RDP', 'RSH',
        'RKR', 'GRM', 'GUR', 'KOT', 'JDH',
        'UDR', 'AJM', 'BKR', 'BLW', 'SIK',
        'SGG', 'ALW', 'HIS', 'ROH', 'SON',
        'PPT', 'KRN', 'AMB', 'YMG', 'KRK',
        'JMU', 'ATQ', 'SRG', 'PTK', 'NBZ',
        'LXR', 'FAR', 'SDL', 'JAI', 'ALL',
        'KNP', 'VNS', 'MAU', 'MRZ', 'AYP',
        'ALG', 'DEO', 'JNP', 'JHS', 'RBR',
        'BTD', 'CAR', 'JLD', 'LDH', 'LUD',
        'PTL', 'RUP', 'SHM', 'MPR', 'MHP',
        'NDL'
    ]
    ALLOWED_DCS_SET_LOWER = set(dc.lower() for dc in ALLOWED_SOURCE_DCS)

from core.stream_engine import (
    XmlSheetWriter,
    assemble_stream_workbook,
    open_stream_reader,
    get_sheet_names,
    ColumnFinder
)

log = logging.getLogger("ei_stream_server.weekly_scm_tat")


def aging_sort_key(bucket_str: str) -> Tuple[int, int, str]:
    """Sort key for aging buckets ensuring '<24Hrs', '24Hrs', '48Hrs', '>48 Hrs', etc. sort logically."""
    s = str(bucket_str).strip()
    nums = re.findall(r'\d+', s)
    val = int(nums[0]) if nums else 9999
    if '<' in s:
        return (val, 0, s)
    elif '>' in s or '+' in s:
        return (val, 2, s)
    else:
        return (val, 1, s)


def normalize_aging_bucket(raw_aging: Any) -> str:
    """Normalize raw aging values. Maps empty/dash to '<24Hrs' and handles HTML entities."""
    if raw_aging is None:
        return "<24Hrs"
    s = html.unescape(str(raw_aging)).strip()
    if not s or s in ('-', 'none', 'null', 'None'):
        return "<24Hrs"
    s_clean = s.replace(" ", "")
    if s_clean.lower() in ('<24hrs', '<24hr', '<24h', '&lt;24hrs', '&lt;24hr', '&lt;24h'):
        return "<24Hrs"
    if s_clean.lower() in ('24hrs', '24hr', '24h'):
        return "24Hrs"
    if s_clean.lower() in ('48hrs', '48hr', '48h'):
        return "48Hrs"
    if s_clean.lower() in ('>48hrs', '>48hr', '>48h', '&gt;48hrs', '&gt;48hr', '&gt;48h', '>48', '&gt;48'):
        return ">48 Hrs"
    return s


def parse_dc_sheet(input_path: Path, target_sheet: str) -> Tuple[Dict[str, List[Any]], Dict[str, List[Any]], Dict[str, List[Any]], List[str]]:
    """
    Dynamically parses the 3 parallel tables in the 'DC' tab:
      - Overall DC
      - Forward DC
      - Reverse DC
    Detects table column positions and date column count dynamically, adapting to growing or shrinking weeks.
    """
    overall_dict = {}
    forward_dict = {}
    reverse_dict = {}
    date_headers = []

    with open_stream_reader(input_path, sheet_name=target_sheet) as (headers, row_iter):
        overall_col = None
        forward_col = None
        reverse_col = None

        for c_idx, val in enumerate(headers):
            v = str(val or '').strip().lower()
            if 'overall' in v:
                overall_col = c_idx
            elif 'forward' in v:
                forward_col = c_idx
            elif 'reverse' in v:
                reverse_col = c_idx

        if overall_col is None:
            overall_col = 0

        date_row = None
        for row in row_iter:
            if row and any(row):
                date_row = row
                break

        if date_row is None:
            return overall_dict, forward_dict, reverse_dict, date_headers

        def get_dates_and_indices(start_c):
            if start_c is None:
                return [], []
            d_idx = []
            d_hdrs = []
            for c in range(start_c + 1, len(date_row)):
                cell_val = str(date_row[c] or '').strip()
                if cell_val.lower() == 'grand total' or not cell_val:
                    break
                d_idx.append(c)
                d_hdrs.append(cell_val)
            return d_hdrs, d_idx

        o_headers, o_indices = get_dates_and_indices(overall_col)
        f_headers, f_indices = get_dates_and_indices(forward_col)
        r_headers, r_indices = get_dates_and_indices(reverse_col)

        date_headers = o_headers or f_headers or r_headers

        # Process subsequent DC rows
        for row in row_iter:
            if not row or not any(row):
                continue

            # Table 1: Overall DC
            if overall_col is not None and len(row) > overall_col and row[overall_col]:
                dc_code = str(row[overall_col]).strip().upper()
                if dc_code and dc_code.lower() != 'grand total':
                    vals = [row[c] if len(row) > c else None for c in o_indices]
                    overall_dict[dc_code] = vals

            # Table 2: Forward DC
            if forward_col is not None and len(row) > forward_col and row[forward_col]:
                dc_code = str(row[forward_col]).strip().upper()
                if dc_code and dc_code.lower() != 'grand total':
                    vals = [row[c] if len(row) > c else None for c in f_indices]
                    forward_dict[dc_code] = vals

            # Table 3: Reverse DC
            if reverse_col is not None and len(row) > reverse_col and row[reverse_col]:
                dc_code = str(row[reverse_col]).strip().upper()
                if dc_code and dc_code.lower() != 'grand total':
                    vals = [row[c] if len(row) > c else None for c in r_indices]
                    reverse_dict[dc_code] = vals

    return overall_dict, forward_dict, reverse_dict, date_headers


def build_summary_sheet(ws, overall_dict: dict, forward_dict: dict, reverse_dict: dict, date_headers: list):
    """
    Builds the formatted 'summary' tab with 3 side-by-side tables filtered to configured DCs.
    Dynamically adjusts table widths, gaps, and formulas based on date count.
    Applies Green/Yellow/Red conditional color thresholds:
      - Green  (> 95%)        : Fill #C6EFCE, Font #006100
      - Yellow (90% to 95%)   : Fill #FFEB9C, Font #9C5700
      - Red    (< 90%)        : Fill #FFC7CE, Font #9C0006
    """
    ws.sheet_view.showGridLines = True
    num_dates = len(date_headers)
    if num_dates == 0:
        num_dates = 7
        date_headers = [f"Col{i+1}" for i in range(num_dates)]

    font_family = "Segoe UI"
    banner_fill_1 = PatternFill(start_color="1E3A8A", end_color="1E3A8A", fill_type="solid")  # Navy
    banner_fill_2 = PatternFill(start_color="0D9488", end_color="0D9488", fill_type="solid")  # Teal
    banner_fill_3 = PatternFill(start_color="7C3AED", end_color="7C3AED", fill_type="solid")  # Purple
    subhead_fill  = PatternFill(start_color="334155", end_color="334155", fill_type="solid")  # Slate
    zebra_fill    = PatternFill(start_color="F8FAFC", end_color="F8FAFC", fill_type="solid")
    total_fill    = PatternFill(start_color="E2E8F0", end_color="E2E8F0", fill_type="solid")

    green_fill  = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")
    green_font  = Font(name=font_family, size=9, bold=True, color="006100")
    yellow_fill = PatternFill(start_color="FFEB9C", end_color="FFEB9C", fill_type="solid")
    yellow_font = Font(name=font_family, size=9, bold=True, color="9C5700")
    red_fill    = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
    red_font    = Font(name=font_family, size=9, bold=True, color="9C0006")

    font_banner = Font(name=font_family, size=11, bold=True, color="FFFFFF")
    font_subhead = Font(name=font_family, size=9, bold=True, color="FFFFFF")
    font_bold = Font(name=font_family, size=9, bold=True, color="0F172A")
    font_normal = Font(name=font_family, size=9, color="334155")
    font_dc = Font(name=font_family, size=9, bold=True, color="1E293B")

    thin_border_side = Side(style="thin", color="CBD5E1")
    cell_border = Border(left=thin_border_side, right=thin_border_side, top=thin_border_side, bottom=thin_border_side)
    double_bottom_side = Side(style="double", color="64748B")
    total_border = Border(left=thin_border_side, right=thin_border_side, top=thin_border_side, bottom=double_bottom_side)

    center_align = Alignment(horizontal="center", vertical="center")
    right_align = Alignment(horizontal="right", vertical="center")

    def get_cell_style(val, is_zebra):
        if isinstance(val, (int, float)):
            if val > 0.95:
                return green_fill, green_font
            elif val >= 0.90:
                return yellow_fill, yellow_font
            else:
                return red_fill, red_font
        return (zebra_fill if is_zebra else None), font_normal

    # Table start/end positions based on num_dates
    t1_start = 1
    t1_end = t1_start + num_dates
    gap1_col = t1_end + 1
    
    t2_start = gap1_col + 1
    t2_end = t2_start + num_dates
    gap2_col = t2_end + 1
    
    t3_start = gap2_col + 1
    t3_end = t3_start + num_dates

    # Row 1: Top Section Banners
    banners = [
        (t1_start, t1_end, "Overall DC - TAT Adherence", banner_fill_1),
        (t2_start, t2_end, "Forward DC - TAT Adherence", banner_fill_2),
        (t3_start, t3_end, "Reverse DC - TAT Adherence", banner_fill_3),
    ]
    for c_st, c_en, label, b_fill in banners:
        ws.cell(1, c_st, label).font = font_banner
        ws.merge_cells(start_row=1, start_column=c_st, end_row=1, end_column=c_en)
        for col in range(c_st, c_en + 1):
            c = ws.cell(1, col)
            c.fill = b_fill
            c.alignment = center_align

    ws.row_dimensions[1].height = 24

    # Row 2: Sub-headers
    sub_configs = [
        (t1_start, "Source DC", date_headers),
        (t2_start, "Forward DC", date_headers),
        (t3_start, "Reverse DC", date_headers),
    ]
    for start_col, label, d_headers in sub_configs:
        c = ws.cell(2, start_col, label)
        c.font = font_subhead
        c.fill = subhead_fill
        c.alignment = center_align
        c.border = cell_border
        for i, dh in enumerate(d_headers):
            c_dh = ws.cell(2, start_col + 1 + i, dh)
            c_dh.font = font_subhead
            c_dh.fill = subhead_fill
            c_dh.alignment = center_align
            c_dh.border = cell_border

    ws.row_dimensions[2].height = 20

    active_dcs = sorted([dc for dc in ALLOWED_SOURCE_DCS if dc.upper() != 'ALL'])
    curr_row = 3

    stats_overall = [[0.0, 0] for _ in range(num_dates)]
    stats_fwd = [[0.0, 0] for _ in range(num_dates)]
    stats_rev = [[0.0, 0] for _ in range(num_dates)]

    for idx, dc in enumerate(active_dcs):
        is_zebra = (idx % 2 == 1)
        row_fill = zebra_fill if is_zebra else None
        ws.row_dimensions[curr_row].height = 18

        tables_data = [
            (t1_start, overall_dict.get(dc, [None] * num_dates), stats_overall),
            (t2_start, forward_dict.get(dc, [None] * num_dates), stats_fwd),
            (t3_start, reverse_dict.get(dc, [None] * num_dates), stats_rev),
        ]
        for st_col, vals, stat_arr in tables_data:
            c_dc = ws.cell(curr_row, st_col, dc)
            c_dc.font = font_dc
            c_dc.alignment = center_align
            c_dc.border = cell_border
            if row_fill:
                c_dc.fill = row_fill

            for i in range(num_dates):
                v = vals[i] if len(vals) > i else None
                cell = ws.cell(curr_row, st_col + 1 + i)
                cell.border = cell_border
                f_fill, f_font = get_cell_style(v, is_zebra)
                if f_fill:
                    cell.fill = f_fill
                cell.font = f_font
                if isinstance(v, (int, float)):
                    cell.value = v
                    cell.number_format = "0.0%"
                    cell.alignment = right_align
                    stat_arr[i][0] += v
                    stat_arr[i][1] += 1
                else:
                    cell.value = "-"
                    cell.alignment = center_align

        curr_row += 1

    # Summary Row: Server DCs Avg
    ws.row_dimensions[curr_row].height = 20
    summary_sections = [
        (t1_start, "Server DCs Avg", stats_overall),
        (t2_start, "Server DCs Avg", stats_fwd),
        (t3_start, "Server DCs Avg", stats_rev),
    ]
    for start_col, lbl, stat_arr in summary_sections:
        c_tot = ws.cell(curr_row, start_col, lbl)
        c_tot.font = font_bold
        c_tot.fill = total_fill
        c_tot.alignment = center_align
        c_tot.border = total_border
        for i, (tot_sum, cnt) in enumerate(stat_arr):
            cell = ws.cell(curr_row, start_col + 1 + i)
            cell.border = total_border
            if cnt > 0:
                avg_val = round(tot_sum / cnt, 3)
                cell.value = avg_val
                cell.number_format = "0.0%"
                cell.alignment = right_align
                f_fill, f_font = get_cell_style(avg_val, False)
                if f_fill:
                    cell.fill = f_fill
                cell.font = f_font
            else:
                cell.value = "-"
                cell.font = font_bold
                cell.fill = total_fill
                cell.alignment = center_align

    # Set column widths dynamically
    for col in range(1, t3_end + 1):
        letter = get_column_letter(col)
        if col in (gap1_col, gap2_col):
            ws.column_dimensions[letter].width = 3
        elif col in (t1_start, t2_start, t3_start):
            ws.column_dimensions[letter].width = 15
        else:
            ws.column_dimensions[letter].width = 12


def build_task_summary_sheet_from_metrics(ws, metrics: dict):
    """
    Builds the 'today's task summary' tab:
      - 3 KPI Cards: OPEN TASKS, FORWARD FLOW, REVERSE FLOW
      - Dynamic Side-by-Side Arrangement:
          * Left         : DC x Aging Matrix (DC, Aging buckets, Total Open)
          * Gap 1        : 1 Blank separator column (width 4)
          * Center/Right : L4 Issue Category Breakdown
          * Gap 2        : 1 Blank separator column (width 4)
          * Right        : L5 Root Reason Breakdown
    """
    ws.sheet_view.showGridLines = True
    font_family = "Segoe UI"

    total_open_tasks = metrics["total_open_tasks"]
    forward_count = metrics["forward_count"]
    reverse_count = metrics["reverse_count"]
    dc_aging_map = metrics["dc_aging_map"]
    sorted_agings = metrics["sorted_agings"]
    l4_counts = metrics["l4_counts"]
    l5_counts = metrics["l5_counts"]

    card_header_fill = PatternFill(start_color="1E3A8A", end_color="1E3A8A", fill_type="solid")
    card_body_fill = PatternFill(start_color="EFF6FF", end_color="EFF6FF", fill_type="solid")
    table_header_fill = PatternFill(start_color="334155", end_color="334155", fill_type="solid")
    table_total_fill = PatternFill(start_color="E2E8F0", end_color="E2E8F0", fill_type="solid")
    zebra_fill = PatternFill(start_color="F8FAFC", end_color="F8FAFC", fill_type="solid")

    font_card_title = Font(name=font_family, size=9, bold=True, color="FFFFFF")
    font_card_num = Font(name=font_family, size=16, bold=True, color="1E3A8A")
    font_th = Font(name=font_family, size=9, bold=True, color="FFFFFF")
    font_bold = Font(name=font_family, size=9, bold=True, color="0F172A")
    font_normal = Font(name=font_family, size=9, color="334155")
    font_sec = Font(name=font_family, size=11, bold=True, color="1E293B")

    thin_border_side = Side(style="thin", color="CBD5E1")
    cell_border = Border(left=thin_border_side, right=thin_border_side, top=thin_border_side, bottom=thin_border_side)
    double_bottom_side = Side(style="double", color="64748B")
    total_border = Border(left=thin_border_side, right=thin_border_side, top=thin_border_side, bottom=double_bottom_side)

    center_align = Alignment(horizontal="center", vertical="center")
    right_align = Alignment(horizontal="right", vertical="center")
    left_align = Alignment(horizontal="left", vertical="center")

    # ========================================================
    # SECTION 1: 3 KPI Cards (OPEN TASKS, FORWARD, REVERSE)
    # ========================================================
    kpi_cards = [
        ("OPEN TASKS", total_open_tasks, 1, 3),
        ("FORWARD FLOW", forward_count, 4, 6),
        ("REVERSE FLOW", reverse_count, 7, 9),
    ]

    for title, val, c_start, c_end in kpi_cards:
        ws.merge_cells(start_row=2, start_column=c_start, end_row=2, end_column=c_end)
        cell_t = ws.cell(2, c_start, title)
        cell_t.font = font_card_title
        cell_t.fill = card_header_fill
        cell_t.alignment = center_align
        for col_i in range(c_start, c_end + 1):
            ws.cell(2, col_i).border = cell_border
            ws.cell(2, col_i).fill = card_header_fill

        ws.merge_cells(start_row=3, start_column=c_start, end_row=4, end_column=c_end)
        cell_v = ws.cell(3, c_start, f"{val:,}")
        cell_v.font = font_card_num
        cell_v.fill = card_body_fill
        cell_v.alignment = center_align
        for r_card in (3, 4):
            for col_i in range(c_start, c_end + 1):
                ws.cell(r_card, col_i).border = cell_border
                ws.cell(r_card, col_i).fill = card_body_fill

    ws.row_dimensions[2].height = 18
    ws.row_dimensions[3].height = 18
    ws.row_dimensions[4].height = 18

    # ========================================================
    # SECTION 2: DYNAMIC SIDE-BY-SIDE TABLES (Starting Row 6 / 7)
    # ========================================================
    matrix_headers = ["Source DC"] + [f"Aging: {b}" for b in sorted_agings] + ["Total Open"]
    t1_start = 1
    t1_end = t1_start + len(matrix_headers) - 1
    gap1_col = t1_end + 1

    l4_start = gap1_col + 1
    l4_end = l4_start + 2
    gap2_col = l4_end + 1

    l5_start = gap2_col + 1
    l5_end = l5_start + 2

    # Row 6: Section Headers
    ws.cell(6, t1_start, "Open Tasks by DC & Aging").font = font_sec
    ws.cell(6, l4_start, "Top Issue Categories (L4)").font = font_sec
    ws.cell(6, l5_start, "Top Root Reasons (L5)").font = font_sec
    ws.row_dimensions[6].height = 22

    matrix_start_row = 7
    ws.row_dimensions[matrix_start_row].height = 22

    # Table 1 Headers (DC x Aging)
    for col_i, h in enumerate(matrix_headers, t1_start):
        c = ws.cell(matrix_start_row, col_i, h)
        c.font = font_th
        c.fill = table_header_fill
        c.alignment = center_align
        c.border = cell_border

    # Table 2 Headers (L4)
    l4_headers = ["Issue Category (L4)", "Count", "Share %"]
    for col_i, h in enumerate(l4_headers, l4_start):
        c = ws.cell(matrix_start_row, col_i, h)
        c.font = font_th
        c.fill = table_header_fill
        c.alignment = center_align
        c.border = cell_border

    # Table 3 Headers (L5)
    l5_headers = ["Root Reason (L5)", "Count", "Share %"]
    for col_i, h in enumerate(l5_headers, l5_start):
        c = ws.cell(matrix_start_row, col_i, h)
        c.font = font_th
        c.fill = table_header_fill
        c.alignment = center_align
        c.border = cell_border

    # Data Rows
    aging_col_totals = defaultdict(int)
    sorted_dcs_with_open = sorted(dc_aging_map.keys())
    top_l4 = sorted(l4_counts.items(), key=lambda x: -x[1])[:15]
    top_l5 = sorted(l5_counts.items(), key=lambda x: -x[1])[:15]

    num_dc_rows = len(sorted_dcs_with_open)
    max_side_rows = max(len(top_l4), len(top_l5))
    total_data_rows = max(num_dc_rows, max_side_rows)

    for i in range(total_data_rows):
        r_now = matrix_start_row + 1 + i
        ws.row_dimensions[r_now].height = 18
        is_zebra = (i % 2 == 1)
        r_fill = zebra_fill if is_zebra else None

        # 1. Fill DC x Aging Row
        if i < num_dc_rows:
            dc = sorted_dcs_with_open[i]
            c_dc = ws.cell(r_now, t1_start, dc)
            c_dc.font = font_bold
            c_dc.alignment = center_align
            c_dc.border = cell_border
            if r_fill:
                c_dc.fill = r_fill

            dc_tot = dc_aging_map[dc]['__TOTAL__']

            for b_idx, bucket in enumerate(sorted_agings, t1_start + 1):
                cnt = dc_aging_map[dc][bucket]
                aging_col_totals[bucket] += cnt
                cell = ws.cell(r_now, b_idx, cnt if cnt > 0 else "-")
                cell.font = font_normal
                cell.border = cell_border
                if r_fill:
                    cell.fill = r_fill
                cell.alignment = right_align if cnt > 0 else center_align

            # Total Open column
            c_tot = ws.cell(r_now, t1_end, dc_tot)
            c_tot.font = font_bold
            c_tot.alignment = right_align
            c_tot.border = cell_border
            if r_fill:
                c_tot.fill = r_fill

        # 2. Fill L4 Row
        if i < len(top_l4):
            cat, cnt = top_l4[i]
            c_cat = ws.cell(r_now, l4_start, cat)
            c_cnt = ws.cell(r_now, l4_start + 1, cnt)
            c_pct = ws.cell(r_now, l4_start + 2, cnt / total_open_tasks if total_open_tasks else 0)
            c_cat.font, c_cnt.font, c_pct.font = font_normal, font_normal, font_normal
            c_cat.border, c_cnt.border, c_pct.border = cell_border, cell_border, cell_border
            c_cat.alignment, c_cnt.alignment, c_pct.alignment = left_align, right_align, right_align
            c_pct.number_format = "0.0%"
            if r_fill:
                c_cat.fill, c_cnt.fill, c_pct.fill = r_fill, r_fill, r_fill

        # 3. Fill L5 Row
        if i < len(top_l5):
            cat, cnt = top_l5[i]
            c_cat = ws.cell(r_now, l5_start, cat)
            c_cnt = ws.cell(r_now, l5_start + 1, cnt)
            c_pct = ws.cell(r_now, l5_start + 2, cnt / total_open_tasks if total_open_tasks else 0)
            c_cat.font, c_cnt.font, c_pct.font = font_normal, font_normal, font_normal
            c_cat.border, c_cnt.border, c_pct.border = cell_border, cell_border, cell_border
            c_cat.alignment, c_cnt.alignment, c_pct.alignment = left_align, right_align, right_align
            c_pct.number_format = "0.0%"
            if r_fill:
                c_cat.fill, c_cnt.fill, c_pct.fill = r_fill, r_fill, r_fill

    # Total Row for DC x Aging Matrix
    tot_row_idx = matrix_start_row + 1 + num_dc_rows
    ws.row_dimensions[tot_row_idx].height = 20
    c_all = ws.cell(tot_row_idx, t1_start, "Grand Total")
    c_all.font = font_bold
    c_all.fill = table_total_fill
    c_all.alignment = center_align
    c_all.border = total_border

    for b_idx, bucket in enumerate(sorted_agings, t1_start + 1):
        col_sum = aging_col_totals[bucket]
        c_sum = ws.cell(tot_row_idx, b_idx, col_sum)
        c_sum.font = font_bold
        c_sum.fill = table_total_fill
        c_sum.alignment = right_align
        c_sum.border = total_border

    c_grand = ws.cell(tot_row_idx, t1_end, total_open_tasks)
    c_grand.font = font_bold
    c_grand.fill = table_total_fill
    c_grand.alignment = right_align
    c_grand.border = total_border

    # Dynamic column widths
    ws.column_dimensions[get_column_letter(t1_start)].width = 14  # Source DC
    for b_col in range(t1_start + 1, t1_end):
        ws.column_dimensions[get_column_letter(b_col)].width = 14
    ws.column_dimensions[get_column_letter(t1_end)].width = 14    # Total Open

    ws.column_dimensions[get_column_letter(gap1_col)].width = 4   # Gap 1

    ws.column_dimensions[get_column_letter(l4_start)].width = 24     # L4 Name
    ws.column_dimensions[get_column_letter(l4_start + 1)].width = 12 # L4 Count
    ws.column_dimensions[get_column_letter(l4_start + 2)].width = 12 # L4 Share %

    ws.column_dimensions[get_column_letter(gap2_col)].width = 4   # Gap 2

    ws.column_dimensions[get_column_letter(l5_start)].width = 30     # L5 Name
    ws.column_dimensions[get_column_letter(l5_start + 1)].width = 12 # L5 Count
    ws.column_dimensions[get_column_letter(l5_start + 2)].width = 12 # L5 Share %


def build_hub_l5_matrix_tab(ws, hub_metrics: dict):
    """
    Builds Tab 5: 'hub l5 summary'
    Dynamic Hub-wise matrix with distinct L5 Root Reasons as column headers.
    """
    ws.sheet_view.showGridLines = True
    font_family = "Segoe UI"

    font_banner = Font(name=font_family, size=11, bold=True, color="FFFFFF")
    font_th = Font(name=font_family, size=9, bold=True, color="FFFFFF")
    font_bold = Font(name=font_family, size=9, bold=True, color="0F172A")
    font_normal = Font(name=font_family, size=9, color="334155")

    banner_fill = PatternFill(start_color="1E3A8A", end_color="1E3A8A", fill_type="solid")  # Navy
    th_fill = PatternFill(start_color="334155", end_color="334155", fill_type="solid")
    zebra_fill = PatternFill(start_color="F8FAFC", end_color="F8FAFC", fill_type="solid")
    total_fill = PatternFill(start_color="E2E8F0", end_color="E2E8F0", fill_type="solid")

    thin_border_side = Side(style="thin", color="CBD5E1")
    cell_border = Border(left=thin_border_side, right=thin_border_side, top=thin_border_side, bottom=thin_border_side)
    double_bottom_side = Side(style="double", color="64748B")
    total_border = Border(left=thin_border_side, right=thin_border_side, top=thin_border_side, bottom=double_bottom_side)

    center_align = Alignment(horizontal="center", vertical="center")
    right_align = Alignment(horizontal="right", vertical="center")

    total_open_tasks = hub_metrics["total_open_tasks"]
    dc_l5_map = hub_metrics["dc_l5_map"]
    l5_col_order = hub_metrics["l5_col_order"]

    sorted_dcs = sorted([dc for dc, l5_dict in dc_l5_map.items() if sum(l5_dict.values()) > 0])
    num_cols = 1 + len(l5_col_order) + 1

    # Row 1: Banner Header
    ws.cell(1, 1, "HUB-WISE ROOT REASON MATRIX (L5)").font = font_banner
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=num_cols)
    for c in range(1, num_cols + 1):
        ws.cell(1, c).fill = banner_fill
        ws.cell(1, c).alignment = center_align
    ws.row_dimensions[1].height = 24

    # Row 2: Column Headers
    headers = ["Source DC"] + l5_col_order + ["Total Open"]
    ws.row_dimensions[2].height = 22
    for col_i, h in enumerate(headers, 1):
        c = ws.cell(2, col_i, h)
        c.font = font_th
        c.fill = th_fill
        c.alignment = center_align
        c.border = cell_border

    # Data Rows
    l5_col_totals = defaultdict(int)
    curr_row = 3
    for idx, dc in enumerate(sorted_dcs):
        is_zebra = (idx % 2 == 1)
        r_fill = zebra_fill if is_zebra else None
        ws.row_dimensions[curr_row].height = 18

        c_dc = ws.cell(curr_row, 1, dc)
        c_dc.font = font_bold
        c_dc.alignment = center_align
        c_dc.border = cell_border
        if r_fill:
            c_dc.fill = r_fill

        dc_row_total = 0
        for col_offset, l5_cat in enumerate(l5_col_order, 2):
            cnt = dc_l5_map[dc].get(l5_cat, 0)
            dc_row_total += cnt
            l5_col_totals[l5_cat] += cnt

            cell = ws.cell(curr_row, col_offset, cnt if cnt > 0 else "-")
            cell.font = font_normal
            cell.border = cell_border
            if r_fill:
                cell.fill = r_fill
            cell.alignment = right_align if cnt > 0 else center_align

        c_tot = ws.cell(curr_row, num_cols, dc_row_total)
        c_tot.font = font_bold
        c_tot.alignment = right_align
        c_tot.border = cell_border
        if r_fill:
            c_tot.fill = r_fill

        curr_row += 1

    # Grand Total Row
    tot_row = curr_row
    ws.row_dimensions[tot_row].height = 20
    c_gt = ws.cell(tot_row, 1, "Grand Total")
    c_gt.font = font_bold
    c_gt.fill = total_fill
    c_gt.alignment = center_align
    c_gt.border = total_border

    for col_offset, l5_cat in enumerate(l5_col_order, 2):
        col_sum = l5_col_totals[l5_cat]
        c_sum = ws.cell(tot_row, col_offset, col_sum)
        c_sum.font = font_bold
        c_sum.fill = total_fill
        c_sum.alignment = right_align
        c_sum.border = total_border

    c_all = ws.cell(tot_row, num_cols, total_open_tasks)
    c_all.font = font_bold
    c_all.fill = total_fill
    c_all.alignment = right_align
    c_all.border = total_border

    # Column Widths
    ws.column_dimensions["A"].width = 14
    for col_idx in range(2, num_cols):
        let = get_column_letter(col_idx)
        header_len = len(str(headers[col_idx - 1]))
        ws.column_dimensions[let].width = max(14, min(header_len + 3, 30))
    ws.column_dimensions[get_column_letter(num_cols)].width = 14


def build_hub_l4_l5_breakdown_tab(ws, hub_metrics: dict):
    """
    Builds Tab 6: 'hub l4 l5 breakdown'
    Columns: Source DC, Issue Category (L4), Root Reason (L5), Open Tasks, Share of Hub %, Total Share %
    Total Share % formatted to 4 decimal places ('0.0000%').
    """
    ws.sheet_view.showGridLines = True
    font_family = "Segoe UI"

    font_banner = Font(name=font_family, size=11, bold=True, color="FFFFFF")
    font_th = Font(name=font_family, size=9, bold=True, color="FFFFFF")
    font_bold = Font(name=font_family, size=9, bold=True, color="0F172A")
    font_normal = Font(name=font_family, size=9, color="334155")

    banner_fill = PatternFill(start_color="0D9488", end_color="0D9488", fill_type="solid")  # Teal
    th_fill = PatternFill(start_color="334155", end_color="334155", fill_type="solid")
    zebra_fill = PatternFill(start_color="F8FAFC", end_color="F8FAFC", fill_type="solid")

    thin_border_side = Side(style="thin", color="CBD5E1")
    cell_border = Border(left=thin_border_side, right=thin_border_side, top=thin_border_side, bottom=thin_border_side)

    center_align = Alignment(horizontal="center", vertical="center")
    right_align = Alignment(horizontal="right", vertical="center")
    left_align = Alignment(horizontal="left", vertical="center")

    total_open_tasks = hub_metrics["total_open_tasks"]
    dc_l5_map = hub_metrics["dc_l5_map"]
    dc_l4_l5_list = hub_metrics["dc_l4_l5_list"]

    ws.cell(1, 1, "DETAILED HUB ROOT CAUSE BREAKDOWN (L4 & L5)").font = font_banner
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=6)
    for c in range(1, 7):
        ws.cell(1, c).fill = banner_fill
        ws.cell(1, c).alignment = center_align
    ws.row_dimensions[1].height = 24

    drill_headers = ["Source DC", "Issue Category (L4)", "Root Reason (L5)", "Open Tasks", "Share of Hub %", "Total Share %"]
    ws.row_dimensions[2].height = 20
    for col_i, h in enumerate(drill_headers, 1):
        c = ws.cell(2, col_i, h)
        c.font = font_th
        c.fill = th_fill
        c.alignment = center_align
        c.border = cell_border

    curr_row = 3
    for idx, (dc, l4, l5, cnt) in enumerate(dc_l4_l5_list):
        is_zebra = (idx % 2 == 1)
        r_fill = zebra_fill if is_zebra else None
        ws.row_dimensions[curr_row].height = 18

        hub_total = sum(dc_l5_map[dc].values()) if dc in dc_l5_map else cnt
        hub_share = (cnt / hub_total) if hub_total else 0
        tot_share = (cnt / total_open_tasks) if total_open_tasks else 0

        c_dc = ws.cell(curr_row, 1, dc)
        c_l4 = ws.cell(curr_row, 2, l4)
        c_l5 = ws.cell(curr_row, 3, l5)
        c_cnt = ws.cell(curr_row, 4, cnt)
        c_hshare = ws.cell(curr_row, 5, hub_share)
        c_tshare = ws.cell(curr_row, 6, tot_share)

        c_dc.font = font_bold
        c_l4.font = font_normal
        c_l5.font = font_normal
        c_cnt.font = font_normal
        c_hshare.font = font_normal
        c_tshare.font = font_normal

        c_dc.alignment = center_align
        c_l4.alignment = left_align
        c_l5.alignment = left_align
        c_cnt.alignment = right_align
        c_hshare.alignment = right_align
        c_tshare.alignment = right_align

        c_hshare.number_format = "0.0%"
        c_tshare.number_format = "0.0000%"

        for c_cell in (c_dc, c_l4, c_l5, c_cnt, c_hshare, c_tshare):
            c_cell.border = cell_border
            if r_fill:
                c_cell.fill = r_fill

        curr_row += 1

    ws.column_dimensions["A"].width = 14  # Source DC
    ws.column_dimensions["B"].width = 24  # L4
    ws.column_dimensions["C"].width = 32  # L5
    ws.column_dimensions["D"].width = 14  # Count
    ws.column_dimensions["E"].width = 16  # Share of Hub %
    ws.column_dimensions["F"].width = 16  # Total Share %


def _post_process_xml_entities(xlsx_path: Path):
    """
    Sanitizes XML entries in the generated XLSX archive:
    Replaces '&gt;' with literal '>' across XML worksheet entries.
    Google Drive / Google Sheets has a known importer defect where '&gt;' in cell
    strings is not unescaped, causing cells to display literal 'Aging: &gt;48 Hrs'.
    In standard XML, '>' is completely valid inside character data without escaping.
    """
    temp_path = xlsx_path.with_suffix(".tmp_sanitized.xlsx")
    try:
        with zipfile.ZipFile(xlsx_path, 'r') as z_in, zipfile.ZipFile(temp_path, 'w', compression=zipfile.ZIP_DEFLATED) as z_out:
            for item in z_in.infolist():
                data = z_in.read(item.filename)
                if item.filename.endswith('.xml') and b'&gt;' in data:
                    data = data.replace(b'&gt;', b'>')
                z_out.writestr(item, data)
        temp_path.replace(xlsx_path)
    except Exception as e:
        log.warning(f"Could not post-process XML entities in {xlsx_path.name}: {e}")
        if temp_path.exists():
            try:
                temp_path.unlink()
            except Exception:
                pass


def generate_weekly_scm_tat_report(input_file: Path, output_file: Path):
    """
    Main Zero-Memory Streaming Entrypoint for Weekly SCM TAT Report.
    Processes 10MB - 40MB+ workbooks with O(1) memory footprint (< 35MB RAM).
    """
    input_path = Path(input_file).resolve()
    output_path = Path(output_file).resolve()

    if not input_path.exists():
        raise FileNotFoundError(f"Input file not found: {input_path}")

    log.info(f"Loading input workbook for Weekly SCM TAT Report (Streaming): {input_path.name}")

    all_sheets = get_sheet_names(input_path)
    sheet_map = {name.lower().strip(): name for name in all_sheets}

    # Locate 'DC' sheet
    dc_sheet_name = None
    for cand in ['dc', 'dc_view', 'dc summary']:
        if cand in sheet_map:
            dc_sheet_name = sheet_map[cand]
            break

    overall_dict, forward_dict, reverse_dict, date_headers = {}, {}, {}, []
    if dc_sheet_name:
        log.info(f"Extracting DC adherence tables from sheet: '{dc_sheet_name}'")
        overall_dict, forward_dict, reverse_dict, date_headers = parse_dc_sheet(input_path, dc_sheet_name)
    else:
        log.warning("No 'DC' sheet found in input workbook.")

    # Locate 'Data' sheet
    data_sheet_name = None
    for cand in ['data', 'raw', 'raw_data', 'tasks', 'sheet1']:
        if cand in sheet_map:
            data_sheet_name = sheet_map[cand]
            break
    if not data_sheet_name and all_sheets:
        data_sheet_name = all_sheets[0]

    log.info(f"Streaming granular data rows from sheet: '{data_sheet_name}'")

    # In-flight aggregators (< 1 MB RAM)
    total_filtered = 0
    total_open_tasks = 0
    forward_count = 0
    reverse_count = 0
    aging_buckets = set()
    dc_aging_map = defaultdict(lambda: defaultdict(int))
    l4_counts = defaultdict(int)
    l5_counts = defaultdict(int)
    dc_l5_map = defaultdict(lambda: defaultdict(int))
    dc_l4_l5_counter = Counter()

    with open_stream_reader(input_path, sheet_name=data_sheet_name) as (headers, row_iter):
        if not headers:
            raise ValueError(f"Sheet '{data_sheet_name}' is empty or invalid.")

        cf = ColumnFinder(headers, {
            'sdc': ['source dc', 'source_dc', 'dc', 'sourcedccode', 'sourcedcname', 'origin'],
            'status': ['status_status', 'status', 'task_status', 'attempt_status', 'state'],
            'aging': ['aging', 'agingdays', 'agebucket', 'agingbucket'],
            'attr': ['attribute', 'flow', 'direction', 'movement', 'type'],
            'l4': ['l4_name', 'l4', 'issue_category', 'issue_category_l4'],
            'l5': ['l5_name', 'l5', 'root_reason', 'root_reason_l5', 'root_cause']
        })

        sdc_idx = cf.get('sdc', 26)
        status_idx = cf.get('status', 4)
        aging_idx = cf.get('aging', 34)
        attr_idx = cf.get('attr', 23)
        l4_idx = cf.get('l4', 11)
        l5_idx = cf.get('l5', 12)

        # Zero-DOM Disk XML streaming writers
        raw_writer = XmlSheetWriter("raw data", headers)
        tasks_writer = XmlSheetWriter("todays tasks", headers)

        with raw_writer, tasks_writer:
            for row in row_iter:
                if not row or len(row) <= sdc_idx:
                    continue
                raw_sdc = row[sdc_idx]
                if raw_sdc is None:
                    continue
                dc_upper = normalize_dc_code(raw_sdc)

                if is_allowed_dc(dc_upper) and dc_upper != 'ALL':
                    total_filtered += 1
                    r_out = list(row)
                    r_out[sdc_idx] = dc_upper
                    raw_writer.write_row(r_out)

                    stat_val = str(row[status_idx] or '').strip().lower() if len(row) > status_idx else ''
                    if stat_val != 'closed':
                        total_open_tasks += 1
                        tasks_writer.write_row(r_out)
                        raw_ag = row[aging_idx] if len(row) > aging_idx else None
                        ag = normalize_aging_bucket(raw_ag)
                        att = str(row[attr_idx] or '').strip().lower() if len(row) > attr_idx else ''
                        l4 = str(row[l4_idx] or 'Unknown').strip() if len(row) > l4_idx else 'Unknown'
                        l5 = str(row[l5_idx] or 'Unknown').strip() if len(row) > l5_idx else 'Unknown'

                        if att == 'forward':
                            forward_count += 1
                        elif att == 'reverse':
                            reverse_count += 1

                        aging_buckets.add(ag)
                        dc_aging_map[dc_upper][ag] += 1
                        dc_aging_map[dc_upper]['__TOTAL__'] += 1
                        if l4 and l4 != 'None':
                            l4_counts[l4] += 1
                        if l5 and l5 != 'None':
                            l5_counts[l5] += 1
                            dc_l5_map[dc_upper][l5] += 1
                            dc_l4_l5_counter[(dc_upper, l4, l5)] += 1

    log.info(f"Filtered {total_filtered:,} matching rows ({total_open_tasks:,} open tasks).")

    metrics = {
        "total_open_tasks": total_open_tasks,
        "forward_count": forward_count,
        "reverse_count": reverse_count,
        "dc_aging_map": dc_aging_map,
        "sorted_agings": sorted(aging_buckets, key=aging_sort_key),
        "l4_counts": l4_counts,
        "l5_counts": l5_counts
    }

    l5_col_order = [cat for cat, _ in sorted(l5_counts.items(), key=lambda x: x[1], reverse=True)]
    sorted_l4_l5 = sorted(
        [(dc, l4, l5, cnt) for (dc, l4, l5), cnt in dc_l4_l5_counter.items()],
        key=lambda x: (x[0], -x[3], x[1], x[2])
    )
    hub_metrics = {
        "total_open_tasks": total_open_tasks,
        "dc_l5_map": dc_l5_map,
        "l5_col_order": l5_col_order,
        "dc_l4_l5_list": sorted_l4_l5
    }

    # Construct OpenPyXL micro-workbook (< 2 MB RAM) with 6 sheets in exact tab order
    out_wb = openpyxl.Workbook()

    # Sheet 1: summary
    ws_sum = out_wb.active
    ws_sum.title = "summary"
    build_summary_sheet(ws_sum, overall_dict, forward_dict, reverse_dict, date_headers)

    # Sheet 2: raw data (placeholder replaced by assemble_stream_workbook)
    out_wb.create_sheet(title="raw data")

    # Sheet 3: todays tasks (placeholder replaced by assemble_stream_workbook)
    out_wb.create_sheet(title="todays tasks")

    # Sheet 4: today's task summary
    ws_task_sum = out_wb.create_sheet(title="today's task summary")
    build_task_summary_sheet_from_metrics(ws_task_sum, metrics)

    # Sheet 5: hub l5 summary
    ws_hub_l5 = out_wb.create_sheet(title="hub l5 summary")
    build_hub_l5_matrix_tab(ws_hub_l5, hub_metrics)

    # Sheet 6: hub l4 l5 breakdown
    ws_hub_breakdown = out_wb.create_sheet(title="hub l4 l5 breakdown")
    build_hub_l4_l5_breakdown_tab(ws_hub_breakdown, hub_metrics)

    # Assemble final .xlsx with zero-memory disk streaming
    assemble_stream_workbook(out_wb, [raw_writer, tasks_writer], output_path)

    # Post-process XML entities to prevent Google Sheets from rendering &gt;
    _post_process_xml_entities(output_path)
    log.info(f"Successfully generated Weekly SCM TAT Report: {output_path.name} (6 tabs assembled)")
