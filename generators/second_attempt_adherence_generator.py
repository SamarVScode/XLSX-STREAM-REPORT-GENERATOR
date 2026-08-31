"""
Second Attempt Adherence Stream Report Generator
================================================
Streams 2nd Attempt Adherence data (.xlsx / .xlsb) for target DCs using dc_config.
Generates a 2-tab Excel output:
  1. 'Summary': Compact side-by-side FWD & REV DC metrics.
  2. 'Raw': Combined & filtered raw records from FWD and REV sheets matching Source_DC.
"""

import sys
import math
import logging
from pathlib import Path
from typing import List, Dict, Any, Union, Tuple
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

from config.dc_config import ALLOWED_DCS_SET
from core.stream_engine import stream_sheet_rows, get_sheet_names

TARGET_DCS = set(dc.upper() for dc in ALLOWED_DCS_SET if dc.upper() != 'ALL')

log = logging.getLogger("ei_stream_server.2nd_attempt_adherence")

def safe_int(val: Any) -> int:
    try:
        if val is None:
            return 0
        f = float(val)
        if math.isnan(f) or math.isinf(f):
            return 0
        return int(f)
    except (ValueError, TypeError):
        return 0

def safe_float(val: Any) -> float:
    try:
        if val is None:
            return 0.0
        f = float(val)
        if math.isnan(f) or math.isinf(f):
            return 0.0
        return f
    except (ValueError, TypeError):
        return 0.0

def get_adherence_color_styles(pct: float) -> Tuple[PatternFill, Font]:
    if pct < 0.85:
        fill = PatternFill(start_color="FEE2E2", end_color="FEE2E2", fill_type="solid")
        font = Font(name="Calibri", size=10, bold=True, color="991B1B")
    elif pct <= 0.95:
        fill = PatternFill(start_color="FEF3C7", end_color="FEF3C7", fill_type="solid")
        font = Font(name="Calibri", size=10, bold=True, color="92400E")
    else:
        fill = PatternFill(start_color="D1FAE5", end_color="D1FAE5", fill_type="solid")
        font = Font(name="Calibri", size=10, bold=True, color="065F46")
    return fill, font

def generate_second_attempt_adherence_report(input_path: Union[str, Path], output_path: Union[str, Path]) -> str:
    path = Path(input_path)
    output_path = Path(output_path)
    log.info(f"Stream processing Second Attempt Adherence report: {path.name}")

    if not path.exists():
        raise FileNotFoundError(f"Input file not found: {path}")

    sheet_names = get_sheet_names(path)
    sheet_map = {s.lower(): s for s in sheet_names}

    summary_rows = []
    if "summary" in sheet_map:
        summary_rows = list(stream_sheet_rows(path, sheet_name=sheet_map["summary"]))

    fwd_dict = {}
    rev_dict = {}

    for r in summary_rows[2:]:
        # FWD (Cols 0-4)
        if len(r) >= 5 and r[0] is not None:
            dc = str(r[0]).strip().upper()
            if dc in TARGET_DCS:
                non_adh = safe_int(r[1])
                adh = safe_int(r[2])
                total = safe_int(r[3]) if len(r) > 3 and r[3] is not None else (non_adh + adh)
                pct = safe_float(r[4]) if (len(r) > 4 and r[4] is not None and not math.isnan(safe_float(r[4]))) else (adh / total if total > 0 else 0.0)
                fwd_dict[dc] = {
                    "non_adherence": non_adh,
                    "adherence": adh,
                    "grand_total": total,
                    "adherence_pct": pct
                }

        # REV (Cols 6-10)
        if len(r) >= 11 and r[6] is not None:
            dc = str(r[6]).strip().upper()
            if dc in TARGET_DCS:
                non_adh = safe_int(r[7])
                adh = safe_int(r[8])
                total = safe_int(r[9]) if len(r) > 9 and r[9] is not None else (non_adh + adh)
                pct = safe_float(r[10]) if (len(r) > 10 and r[10] is not None and not math.isnan(safe_float(r[10]))) else (adh / total if total > 0 else 0.0)
                rev_dict[dc] = {
                    "non_adherence": non_adh,
                    "adherence": adh,
                    "grand_total": total,
                    "adherence_pct": pct
                }

    sorted_dcs = sorted(list(set(fwd_dict.keys()) | set(rev_dict.keys())))

    filtered_raw_records = []
    raw_headers = []

    # Stream FWD Raw
    if "fwd" in sheet_map:
        fwd_iter = stream_sheet_rows(path, sheet_name=sheet_map["fwd"])
        try:
            fwd_h = next(fwd_iter)
            fwd_headers = [str(h).strip() if h is not None else "" for h in fwd_h]
            raw_headers = ["Flow_Type"] + fwd_headers
            
            dc_idx = -1
            for name in ["Source_DC", "source_dc", "DC", "dc"]:
                if name in fwd_headers:
                    dc_idx = fwd_headers.index(name)
                    break

            if dc_idx != -1:
                for row in fwd_iter:
                    if len(row) > dc_idx and row[dc_idx] is not None:
                        dc_val = str(row[dc_idx]).strip().upper()
                        if dc_val in TARGET_DCS:
                            filtered_raw_records.append(["FWD"] + list(row))
        except StopIteration:
            pass

    # Stream REV Raw
    if "rev" in sheet_map:
        rev_iter = stream_sheet_rows(path, sheet_name=sheet_map["rev"])
        try:
            rev_h = next(rev_iter)
            rev_headers = [str(h).strip() if h is not None else "" for h in rev_h]
            if not raw_headers:
                raw_headers = ["Flow_Type"] + rev_headers

            dc_idx = -1
            for name in ["Source_DC", "source_dc", "DC", "dc"]:
                if name in rev_headers:
                    dc_idx = rev_headers.index(name)
                    break

            if dc_idx != -1:
                for row in rev_iter:
                    if len(row) > dc_idx and row[dc_idx] is not None:
                        dc_val = str(row[dc_idx]).strip().upper()
                        if dc_val in TARGET_DCS:
                            filtered_raw_records.append(["REV"] + list(row))
        except StopIteration:
            pass

    out_wb = openpyxl.Workbook()
    
    # --- Tab 1: Summary ---
    ws_sum = out_wb.active
    ws_sum.title = "Summary"
    ws_sum.sheet_view.showGridLines = True

    font_top_header = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    font_fwd_sub = Font(name="Calibri", size=9, bold=True, color="1F497D")
    font_rev_sub = Font(name="Calibri", size=9, bold=True, color="375623")
    font_regular = Font(name="Calibri", size=9)

    fill_fwd_top = PatternFill(start_color="2F5597", end_color="2F5597", fill_type="solid")
    fill_rev_top = PatternFill(start_color="375623", end_color="375623", fill_type="solid")

    fill_fwd_sub = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")
    fill_rev_sub = PatternFill(start_color="E2EFDA", end_color="E2EFDA", fill_type="solid")

    border_thin = Border(
        left=Side(style="thin", color="D9D9D9"),
        right=Side(style="thin", color="D9D9D9"),
        top=Side(style="thin", color="D9D9D9"),
        bottom=Side(style="thin", color="D9D9D9")
    )

    ws_sum.merge_cells("A1:E1")
    fwd_h_cell = ws_sum.cell(row=1, column=1, value="FWD")
    fwd_h_cell.font = font_top_header
    fwd_h_cell.alignment = Alignment(horizontal="center", vertical="center")

    ws_sum.merge_cells("G1:K1")
    rev_h_cell = ws_sum.cell(row=1, column=7, value="REV")
    rev_h_cell.font = font_top_header
    rev_h_cell.alignment = Alignment(horizontal="center", vertical="center")

    for c in range(1, 6):
        ws_sum.cell(row=1, column=c).fill = fill_fwd_top
    for c in range(7, 12):
        ws_sum.cell(row=1, column=c).fill = fill_rev_top

    headers_block_short = ["DC", "Non-Adh", "2nd Adh", "Total", "Adh %"]
    
    for idx, h in enumerate(headers_block_short, 1):
        cell = ws_sum.cell(row=2, column=idx, value=h)
        cell.font = font_fwd_sub
        cell.fill = fill_fwd_sub
        cell.alignment = Alignment(horizontal="center" if idx > 1 else "left", vertical="center")
        cell.border = border_thin

    for idx, h in enumerate(headers_block_short, 7):
        cell = ws_sum.cell(row=2, column=idx, value=h)
        cell.font = font_rev_sub
        cell.fill = fill_rev_sub
        cell.alignment = Alignment(horizontal="center" if idx > 1 else "left", vertical="center")
        cell.border = border_thin

    current_row = 3
    for dc in sorted_dcs:
        fwd_item = fwd_dict.get(dc, {"non_adherence": 0, "adherence": 0, "grand_total": 0, "adherence_pct": 0.0})
        rev_item = rev_dict.get(dc, {"non_adherence": 0, "adherence": 0, "grand_total": 0, "adherence_pct": 0.0})

        ws_sum.cell(row=current_row, column=1, value=dc).font = font_regular
        ws_sum.cell(row=current_row, column=2, value=fwd_item["non_adherence"]).number_format = "#,##0"
        ws_sum.cell(row=current_row, column=3, value=fwd_item["adherence"]).number_format = "#,##0"
        ws_sum.cell(row=current_row, column=4, value=fwd_item["grand_total"]).number_format = "#,##0"
        pct_fwd = ws_sum.cell(row=current_row, column=5, value=fwd_item["adherence_pct"])
        pct_fwd.number_format = "0.0%"
        fwd_fill, fwd_font = get_adherence_color_styles(fwd_item["adherence_pct"])
        pct_fwd.fill = fwd_fill
        pct_fwd.font = fwd_font

        ws_sum.cell(row=current_row, column=7, value=dc).font = font_regular
        ws_sum.cell(row=current_row, column=8, value=rev_item["non_adherence"]).number_format = "#,##0"
        ws_sum.cell(row=current_row, column=9, value=rev_item["adherence"]).number_format = "#,##0"
        ws_sum.cell(row=current_row, column=10, value=rev_item["grand_total"]).number_format = "#,##0"
        pct_rev = ws_sum.cell(row=current_row, column=11, value=rev_item["adherence_pct"])
        pct_rev.number_format = "0.0%"
        rev_fill, rev_font = get_adherence_color_styles(rev_item["adherence_pct"])
        pct_rev.fill = rev_fill
        pct_rev.font = rev_font

        for c in range(1, 6):
            cell = ws_sum.cell(row=current_row, column=c)
            cell.border = border_thin
            if c > 1: cell.alignment = Alignment(horizontal="right")

        for c in range(7, 12):
            cell = ws_sum.cell(row=current_row, column=c)
            cell.border = border_thin
            if c > 1: cell.alignment = Alignment(horizontal="right")

        current_row += 1

    compact_widths = {
        "A": 9,  "B": 10, "C": 10, "D": 9,  "E": 9,
        "F": 3,
        "G": 9,  "H": 10, "I": 10, "J": 9,  "K": 9
    }
    for col_letter, width in compact_widths.items():
        ws_sum.column_dimensions[col_letter].width = width

    # --- Tab 2: Raw ---
    ws_raw = out_wb.create_sheet(title="Raw")
    ws_raw.sheet_view.showGridLines = True

    if raw_headers:
        ws_raw.append(raw_headers)
        font_raw_header = Font(name="Calibri", size=10, bold=True, color="FFFFFF")
        fill_raw_header = PatternFill(start_color="334155", end_color="334155", fill_type="solid")
        for col_num in range(1, len(raw_headers) + 1):
            cell = ws_raw.cell(row=1, column=col_num)
            cell.font = font_raw_header
            cell.fill = fill_raw_header
            cell.alignment = Alignment(horizontal="center", vertical="center")

            col_letter = get_column_letter(col_num)
            h_len = len(str(raw_headers[col_num - 1]))
            ws_raw.column_dimensions[col_letter].width = min(max(h_len + 4, 12), 35)

    for record in filtered_raw_records:
        ws_raw.append(record)

    output_path.parent.mkdir(parents=True, exist_ok=True)
    out_wb.save(output_path)
    log.info(f"Successfully generated 2nd Attempt Adherence report: {output_path}")
    return str(output_path)
