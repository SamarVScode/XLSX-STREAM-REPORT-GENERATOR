#!/usr/bin/env python3
"""
Weekly SCM TAT & HCQ Report Generator Module for ei_stream_server
================================================================
Generates a structured 4-tab workbook for SCM TAT & HCQ:
  1. 'summary'              : Filtered & aligned DC scorecard with Green/Yellow/Red coloring
  2. 'raw data'             : All rows for configured DCs from the 'Data' tab (Streamed to Disk XML)
  3. 'todays tasks'         : Open/pending tasks (status != 'Closed') (Streamed to Disk XML)
  4. "today's task summary" : Executive KPI cards (Open, Forward, Reverse), DC x Aging matrix,
                              and side-by-side L4 & L5 Root Cause breakdowns.

Uses Single-Pass Zero-Memory Streaming Engine (core.stream_engine):
- O(1) Memory Footprint (< 35MB RAM)
- Direct XML disk streaming for massive (10MB - 40MB+) datasets
"""

import sys
import logging
from pathlib import Path
from collections import defaultdict
import datetime
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


def parse_dc_sheet(input_path: Path, target_sheet: str) -> Tuple[Dict[str, List[Any]], Dict[str, List[Any]], Dict[str, List[Any]], List[str]]:
    """
    Parses the 3 parallel tables in the 'DC' tab:
      - Overall DC (Cols A-H / index 0, 1..7)
      - Forward DC (Cols J-Q / index 9, 10..16)
      - Reverse DC (Cols S-Z / index 18, 19..25)
    """
    overall_dict = {}
    forward_dict = {}
    reverse_dict = {}
    date_headers = []

    with open_stream_reader(input_path, sheet_name=target_sheet) as (headers, row_iter):
        # Row 1 is header ('Overall DC', ...)
        # Row 2 contains the date headers in cols 1..7
        r2 = None
        for row in row_iter:
            if row and any(row):
                r2 = row
                break

        if r2 is not None:
            for i in range(1, 8):
                if len(r2) > i and r2[i] is not None:
                    val = str(r2[i]).strip()
                    date_headers.append(val)
                else:
                    date_headers.append(f"Col{i+1}")

        # Process subsequent DC rows
        for row in row_iter:
            if not row or not any(row):
                continue

            # Table 1: Overall DC (Col A=0, B-H=1..7)
            if len(row) > 0 and row[0] is not None:
                dc_code = str(row[0]).strip()
                if dc_code and dc_code.lower() != 'grand total':
                    vals = [row[c] if len(row) > c else None for c in range(1, 8)]
                    overall_dict[dc_code.upper()] = vals

            # Table 2: Forward DC (Col J=9, K-Q=10..16)
            if len(row) > 9 and row[9] is not None:
                dc_code = str(row[9]).strip()
                if dc_code and dc_code.lower() != 'grand total':
                    vals = [row[c] if len(row) > c else None for c in range(10, 17)]
                    forward_dict[dc_code.upper()] = vals

            # Table 3: Reverse DC (Col S=18, T-Z=19..25)
            if len(row) > 18 and row[18] is not None:
                dc_code = str(row[18]).strip()
                if dc_code and dc_code.lower() != 'grand total':
                    vals = [row[c] if len(row) > c else None for c in range(19, 26)]
                    reverse_dict[dc_code.upper()] = vals

    return overall_dict, forward_dict, reverse_dict, date_headers


def build_summary_sheet(ws, overall_dict: dict, forward_dict: dict, reverse_dict: dict, date_headers: list):
    """
    Builds the formatted 'summary' tab with 3 side-by-side tables filtered to configured DCs.
    Applies Green/Yellow/Red conditional color thresholds:
      - Green  (> 95%)        : Fill #C6EFCE, Font #006100
      - Yellow (90% to 95%)   : Fill #FFEB9C, Font #9C5700
      - Red    (< 90%)        : Fill #FFC7CE, Font #9C0006
    """
    ws.sheet_view.showGridLines = True

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

    # Row 1: Top Section Banners
    ws.cell(1, 1, "Overall DC - TAT Adherence").font = font_banner
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=8)
    for col in range(1, 9):
        c = ws.cell(1, col)
        c.fill = banner_fill_1
        c.alignment = center_align

    ws.cell(1, 10, "Forward DC - TAT Adherence").font = font_banner
    ws.merge_cells(start_row=1, start_column=10, end_row=1, end_column=17)
    for col in range(10, 18):
        c = ws.cell(1, col)
        c.fill = banner_fill_2
        c.alignment = center_align

    ws.cell(1, 19, "Reverse DC - TAT Adherence").font = font_banner
    ws.merge_cells(start_row=1, start_column=19, end_row=1, end_column=26)
    for col in range(19, 27):
        c = ws.cell(1, col)
        c.fill = banner_fill_3
        c.alignment = center_align

    ws.row_dimensions[1].height = 24

    # Row 2: Sub-headers
    sub_configs = [
        (1, "Source DC", date_headers),
        (10, "Forward DC", date_headers),
        (19, "Reverse DC", date_headers),
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

    stats_overall = [[0.0, 0] for _ in range(7)]
    stats_fwd = [[0.0, 0] for _ in range(7)]
    stats_rev = [[0.0, 0] for _ in range(7)]

    for idx, dc in enumerate(active_dcs):
        is_zebra = (idx % 2 == 1)
        row_fill = zebra_fill if is_zebra else None
        ws.row_dimensions[curr_row].height = 18

        # Overall Table
        c_dc1 = ws.cell(curr_row, 1, dc)
        c_dc1.font = font_dc
        c_dc1.alignment = center_align
        c_dc1.border = cell_border
        if row_fill:
            c_dc1.fill = row_fill
        vals1 = overall_dict.get(dc, [None] * 7)
        for i, v in enumerate(vals1):
            cell = ws.cell(curr_row, 2 + i)
            cell.border = cell_border
            f_fill, f_font = get_cell_style(v, is_zebra)
            if f_fill:
                cell.fill = f_fill
            cell.font = f_font
            if isinstance(v, (int, float)):
                cell.value = v
                cell.number_format = "0.0%"
                cell.alignment = right_align
                stats_overall[i][0] += v
                stats_overall[i][1] += 1
            else:
                cell.value = "-"
                cell.alignment = center_align

        # Forward Table
        c_dc2 = ws.cell(curr_row, 10, dc)
        c_dc2.font = font_dc
        c_dc2.alignment = center_align
        c_dc2.border = cell_border
        if row_fill:
            c_dc2.fill = row_fill
        vals2 = forward_dict.get(dc, [None] * 7)
        for i, v in enumerate(vals2):
            cell = ws.cell(curr_row, 11 + i)
            cell.border = cell_border
            f_fill, f_font = get_cell_style(v, is_zebra)
            if f_fill:
                cell.fill = f_fill
            cell.font = f_font
            if isinstance(v, (int, float)):
                cell.value = v
                cell.number_format = "0.0%"
                cell.alignment = right_align
                stats_fwd[i][0] += v
                stats_fwd[i][1] += 1
            else:
                cell.value = "-"
                cell.alignment = center_align

        # Reverse Table
        c_dc3 = ws.cell(curr_row, 19, dc)
        c_dc3.font = font_dc
        c_dc3.alignment = center_align
        c_dc3.border = cell_border
        if row_fill:
            c_dc3.fill = row_fill
        vals3 = reverse_dict.get(dc, [None] * 7)
        for i, v in enumerate(vals3):
            cell = ws.cell(curr_row, 20 + i)
            cell.border = cell_border
            f_fill, f_font = get_cell_style(v, is_zebra)
            if f_fill:
                cell.fill = f_fill
            cell.font = f_font
            if isinstance(v, (int, float)):
                cell.value = v
                cell.number_format = "0.0%"
                cell.alignment = right_align
                stats_rev[i][0] += v
                stats_rev[i][1] += 1
            else:
                cell.value = "-"
                cell.alignment = center_align

        curr_row += 1

    # Summary Row: Server DCs Avg
    ws.row_dimensions[curr_row].height = 20
    summary_sections = [
        (1, "Server DCs Avg", stats_overall),
        (10, "Server DCs Avg", stats_fwd),
        (19, "Server DCs Avg", stats_rev),
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

    # Column widths
    for col in range(1, 27):
        letter = get_column_letter(col)
        if col in (9, 18):
            ws.column_dimensions[letter].width = 3
        elif col in (1, 10, 19):
            ws.column_dimensions[letter].width = 15
        else:
            ws.column_dimensions[letter].width = 12


def build_task_summary_sheet_from_metrics(ws, metrics: dict):
    """
    Builds the 'today's task summary' tab:
      - 3 KPI Cards: OPEN TASKS, FORWARD FLOW, REVERSE FLOW
      - Side-by-Side Arrangement:
          * Left  (Cols A-E) : DC x Aging Matrix (DC, Aging buckets, Total Open)
          * Right (Cols G-I) : L4 Issue Category Breakdown
          * Right (Cols K-M) : L5 Root Reason Breakdown
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
    # SECTION 2: SIDE-BY-SIDE TABLES (Starting Row 6 / 7)
    # ========================================================
    ws.cell(6, 1, "Open Tasks by DC & Aging").font = font_sec
    ws.cell(6, 7, "Top Issue Categories (L4)").font = font_sec
    ws.cell(6, 11, "Top Root Reasons (L5)").font = font_sec
    ws.row_dimensions[6].height = 22

    matrix_start_row = 7
    ws.row_dimensions[matrix_start_row].height = 22

    # Table 1 Headers (Cols 1..N) - DC x Aging
    matrix_headers = ["Source DC"] + [f"Aging: {b}" for b in sorted_agings] + ["Total Open"]
    for col_i, h in enumerate(matrix_headers, 1):
        c = ws.cell(matrix_start_row, col_i, h)
        c.font = font_th
        c.fill = table_header_fill
        c.alignment = center_align
        c.border = cell_border

    # Table 2 Headers (Cols 7..9) - L4
    l4_headers = ["Issue Category (L4)", "Count", "Share %"]
    for col_i, h in enumerate(l4_headers, 7):
        c = ws.cell(matrix_start_row, col_i, h)
        c.font = font_th
        c.fill = table_header_fill
        c.alignment = center_align
        c.border = cell_border

    # Table 3 Headers (Cols 11..13) - L5
    l5_headers = ["Root Reason (L5)", "Count", "Share %"]
    for col_i, h in enumerate(l5_headers, 11):
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
            c_dc = ws.cell(r_now, 1, dc)
            c_dc.font = font_bold
            c_dc.alignment = center_align
            c_dc.border = cell_border
            if r_fill:
                c_dc.fill = r_fill

            dc_tot = dc_aging_map[dc]['__TOTAL__']

            for b_idx, bucket in enumerate(sorted_agings, 2):
                cnt = dc_aging_map[dc][bucket]
                aging_col_totals[bucket] += cnt
                cell = ws.cell(r_now, b_idx, cnt if cnt > 0 else "-")
                cell.font = font_normal
                cell.border = cell_border
                if r_fill:
                    cell.fill = r_fill
                cell.alignment = right_align if cnt > 0 else center_align

            # Total Open column
            c_tot = ws.cell(r_now, len(sorted_agings) + 2, dc_tot)
            c_tot.font = font_bold
            c_tot.alignment = right_align
            c_tot.border = cell_border
            if r_fill:
                c_tot.fill = r_fill

        # 2. Fill L4 Row
        if i < len(top_l4):
            cat, cnt = top_l4[i]
            c7 = ws.cell(r_now, 7, cat)
            c8 = ws.cell(r_now, 8, cnt)
            c9 = ws.cell(r_now, 9, cnt / total_open_tasks if total_open_tasks else 0)
            c7.font, c8.font, c9.font = font_normal, font_normal, font_normal
            c7.border, c8.border, c9.border = cell_border, cell_border, cell_border
            c7.alignment, c8.alignment, c9.alignment = left_align, right_align, right_align
            c9.number_format = "0.0%"
            if r_fill:
                c7.fill, c8.fill, c9.fill = r_fill, r_fill, r_fill

        # 3. Fill L5 Row
        if i < len(top_l5):
            cat, cnt = top_l5[i]
            c11 = ws.cell(r_now, 11, cat)
            c12 = ws.cell(r_now, 12, cnt)
            c13 = ws.cell(r_now, 13, cnt / total_open_tasks if total_open_tasks else 0)
            c11.font, c12.font, c13.font = font_normal, font_normal, font_normal
            c11.border, c12.border, c13.border = cell_border, cell_border, cell_border
            c11.alignment, c12.alignment, c13.alignment = left_align, right_align, right_align
            c13.number_format = "0.0%"
            if r_fill:
                c11.fill, c12.fill, c13.fill = r_fill, r_fill, r_fill

    # Total Row for DC x Aging Matrix
    tot_row_idx = matrix_start_row + 1 + num_dc_rows
    ws.row_dimensions[tot_row_idx].height = 20
    c_all = ws.cell(tot_row_idx, 1, "Grand Total")
    c_all.font = font_bold
    c_all.fill = table_total_fill
    c_all.alignment = center_align
    c_all.border = total_border

    for b_idx, bucket in enumerate(sorted_agings, 2):
        col_sum = aging_col_totals[bucket]
        c_sum = ws.cell(tot_row_idx, b_idx, col_sum)
        c_sum.font = font_bold
        c_sum.fill = table_total_fill
        c_sum.alignment = right_align
        c_sum.border = total_border

    c_grand = ws.cell(tot_row_idx, len(sorted_agings) + 2, total_open_tasks)
    c_grand.font = font_bold
    c_grand.fill = table_total_fill
    c_grand.alignment = right_align
    c_grand.border = total_border

    # Set column widths
    ws.column_dimensions["A"].width = 14  # Source DC
    for b_idx in range(2, len(sorted_agings) + 2):
        letter = get_column_letter(b_idx)
        ws.column_dimensions[letter].width = 14
    ws.column_dimensions[get_column_letter(len(sorted_agings) + 2)].width = 14  # Total Open

    ws.column_dimensions["F"].width = 4   # Separator
    ws.column_dimensions["G"].width = 24  # L4 Name
    ws.column_dimensions["H"].width = 12  # L4 Count
    ws.column_dimensions["I"].width = 12  # L4 Share %
    ws.column_dimensions["J"].width = 4   # Separator
    ws.column_dimensions["K"].width = 30  # L5 Name
    ws.column_dimensions["L"].width = 12  # L5 Count
    ws.column_dimensions["M"].width = 12  # L5 Share %


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
                        ag = str(row[aging_idx] or '-').strip() if len(row) > aging_idx else '-'
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

    log.info(f"Filtered {total_filtered:,} matching rows ({total_open_tasks:,} open tasks).")

    metrics = {
        "total_open_tasks": total_open_tasks,
        "forward_count": forward_count,
        "reverse_count": reverse_count,
        "dc_aging_map": dc_aging_map,
        "sorted_agings": sorted(aging_buckets),
        "l4_counts": l4_counts,
        "l5_counts": l5_counts
    }

    # Construct OpenPyXL micro-workbook (< 2 MB RAM) with 4 sheets in exact tab order
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

    # Assemble final .xlsx with zero-memory disk streaming
    assemble_stream_workbook(out_wb, [raw_writer, tasks_writer], output_path)
    log.info(f"Successfully generated Weekly SCM TAT Report: {output_path.name} (4 tabs assembled)")
