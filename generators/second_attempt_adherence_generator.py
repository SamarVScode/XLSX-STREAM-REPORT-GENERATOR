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
from collections import defaultdict
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
    sheet_map = {s.lower().replace(' ', '_'): s for s in sheet_names}

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

    filtered_raw_records = []
    raw_headers = []

    # Stream FWD Raw
    fwd_sheet = None
    for cand in ['fwd', 'forward', 'fwd_raw', 'raw']:
        if cand in sheet_map:
            fwd_sheet = sheet_map[cand]
            break

    if fwd_sheet:
        fwd_iter = stream_sheet_rows(path, sheet_name=fwd_sheet)
        try:
            fwd_h = next(fwd_iter)
            fwd_headers = [str(h).strip() if h is not None else "" for h in fwd_h]
            raw_headers = ["Flow_Type"] + fwd_headers
            
            dc_idx = -1
            adh_idx = -1
            for idx, name in enumerate(fwd_headers):
                n_clean = name.lower().replace('_', ' ')
                if n_clean in ("source dc", "dc", "sdc"):
                    dc_idx = idx
                elif "adherence" in n_clean or "status" in n_clean:
                    adh_idx = idx

            for row in fwd_iter:
                if dc_idx != -1 and len(row) > dc_idx and row[dc_idx] is not None:
                    dc_val = str(row[dc_idx]).strip().upper()
                    if dc_val in TARGET_DCS:
                        filtered_raw_records.append(["FWD"] + list(row))
                        if not fwd_dict or dc_val not in fwd_dict:
                            if dc_val not in fwd_dict:
                                fwd_dict[dc_val] = {"non_adherence": 0, "adherence": 0, "grand_total": 0, "adherence_pct": 0.0}
                            val_str = str(row[adh_idx] or '').strip().lower() if adh_idx != -1 and len(row) > adh_idx else ''
                            if val_str in ('adherence', 'done', 'yes', '1'):
                                fwd_dict[dc_val]["adherence"] += 1
                            else:
                                fwd_dict[dc_val]["non_adherence"] += 1
                            tot = fwd_dict[dc_val]["adherence"] + fwd_dict[dc_val]["non_adherence"]
                            fwd_dict[dc_val]["grand_total"] = tot
                            fwd_dict[dc_val]["adherence_pct"] = fwd_dict[dc_val]["adherence"] / tot if tot > 0 else 0.0
        except StopIteration:
            pass

    # Stream REV Raw
    rev_sheet = None
    for cand in ['rev', 'reverse', 'rev_raw']:
        if cand in sheet_map:
            rev_sheet = sheet_map[cand]
            break

    if rev_sheet:
        rev_iter = stream_sheet_rows(path, sheet_name=rev_sheet)
        try:
            rev_h = next(rev_iter)
            rev_headers = [str(h).strip() if h is not None else "" for h in rev_h]
            if not raw_headers:
                raw_headers = ["Flow_Type"] + rev_headers

            dc_idx = -1
            adh_idx = -1
            for idx, name in enumerate(rev_headers):
                n_clean = name.lower().replace('_', ' ')
                if n_clean in ("source dc", "dc", "sdc"):
                    dc_idx = idx
                elif "adherence" in n_clean or "status" in n_clean:
                    adh_idx = idx

            for row in rev_iter:
                if dc_idx != -1 and len(row) > dc_idx and row[dc_idx] is not None:
                    dc_val = str(row[dc_idx]).strip().upper()
                    if dc_val in TARGET_DCS:
                        filtered_raw_records.append(["REV"] + list(row))
                        if not rev_dict or dc_val not in rev_dict:
                            if dc_val not in rev_dict:
                                rev_dict[dc_val] = {"non_adherence": 0, "adherence": 0, "grand_total": 0, "adherence_pct": 0.0}
                            val_str = str(row[adh_idx] or '').strip().lower() if adh_idx != -1 and len(row) > adh_idx else ''
                            if val_str in ('adherence', 'done', 'yes', '1'):
                                rev_dict[dc_val]["adherence"] += 1
                            else:
                                rev_dict[dc_val]["non_adherence"] += 1
                            tot = rev_dict[dc_val]["adherence"] + rev_dict[dc_val]["non_adherence"]
                            rev_dict[dc_val]["grand_total"] = tot
                            rev_dict[dc_val]["adherence_pct"] = rev_dict[dc_val]["adherence"] / tot if tot > 0 else 0.0
        except StopIteration:
            pass

    sorted_dcs = sorted(list(set(fwd_dict.keys()) | set(rev_dict.keys()) | TARGET_DCS))

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

    fwd_sub_headers = ["Source_DC", "Non-Adherence", "Adherence", "Grand Total", "Adherence %"]
    rev_sub_headers = ["Source_DC", "Non-Adherence", "Adherence", "Grand Total", "Adherence %"]

    for idx, h in enumerate(fwd_sub_headers, start=1):
        cell = ws_sum.cell(row=2, column=idx, value=h)
        cell.font = font_fwd_sub
        cell.fill = fill_fwd_sub
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = border_thin

    for idx, h in enumerate(rev_sub_headers, start=7):
        cell = ws_sum.cell(row=2, column=idx, value=h)
        cell.font = font_rev_sub
        cell.fill = fill_rev_sub
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = border_thin

    row_num = 3
    tot_fwd_non = 0
    tot_fwd_adh = 0
    tot_rev_non = 0
    tot_rev_adh = 0

    for dc in sorted_dcs:
        f_data = fwd_dict.get(dc, {"non_adherence": 0, "adherence": 0, "grand_total": 0, "adherence_pct": 0.0})
        r_data = rev_dict.get(dc, {"non_adherence": 0, "adherence": 0, "grand_total": 0, "adherence_pct": 0.0})

        tot_fwd_non += f_data["non_adherence"]
        tot_fwd_adh += f_data["adherence"]
        tot_rev_non += r_data["non_adherence"]
        tot_rev_adh += r_data["adherence"]

        # FWD Data
        c1 = ws_sum.cell(row=row_num, column=1, value=dc)
        c2 = ws_sum.cell(row=row_num, column=2, value=f_data["non_adherence"])
        c3 = ws_sum.cell(row=row_num, column=3, value=f_data["adherence"])
        c4 = ws_sum.cell(row=row_num, column=4, value=f_data["grand_total"])
        c5 = ws_sum.cell(row=row_num, column=5, value=f_data["adherence_pct"])

        c1.alignment = Alignment(horizontal="center", vertical="center")
        c2.alignment = Alignment(horizontal="right", vertical="center")
        c3.alignment = Alignment(horizontal="right", vertical="center")
        c4.alignment = Alignment(horizontal="right", vertical="center")
        c5.alignment = Alignment(horizontal="right", vertical="center")

        c2.number_format = "#,##0"
        c3.number_format = "#,##0"
        c4.number_format = "#,##0"
        c5.number_format = "0.0%"

        for c in [c1, c2, c3, c4]:
            c.font = font_regular
            c.border = border_thin

        f_fill, f_font = get_adherence_color_styles(f_data["adherence_pct"])
        c5.fill = f_fill
        c5.font = f_font
        c5.border = border_thin

        # REV Data
        c7 = ws_sum.cell(row=row_num, column=7, value=dc)
        c8 = ws_sum.cell(row=row_num, column=8, value=r_data["non_adherence"])
        c9 = ws_sum.cell(row=row_num, column=9, value=r_data["adherence"])
        c10 = ws_sum.cell(row=row_num, column=10, value=r_data["grand_total"])
        c11 = ws_sum.cell(row=row_num, column=11, value=r_data["adherence_pct"])

        c7.alignment = Alignment(horizontal="center", vertical="center")
        c8.alignment = Alignment(horizontal="right", vertical="center")
        c9.alignment = Alignment(horizontal="right", vertical="center")
        c10.alignment = Alignment(horizontal="right", vertical="center")
        c11.alignment = Alignment(horizontal="right", vertical="center")

        c8.number_format = "#,##0"
        c9.number_format = "#,##0"
        c10.number_format = "#,##0"
        c11.number_format = "0.0%"

        for c in [c7, c8, c9, c10]:
            c.font = font_regular
            c.border = border_thin

        r_fill, r_font = get_adherence_color_styles(r_data["adherence_pct"])
        c11.fill = r_fill
        c11.font = r_font
        c11.border = border_thin

        row_num += 1

    # Total Row
    tot_fwd_total = tot_fwd_non + tot_fwd_adh
    tot_fwd_pct = tot_fwd_adh / tot_fwd_total if tot_fwd_total > 0 else 0.0

    tot_rev_total = tot_rev_non + tot_rev_adh
    tot_rev_pct = tot_rev_adh / tot_rev_total if tot_rev_total > 0 else 0.0

    font_total = Font(name="Calibri", size=9, bold=True)

    c1 = ws_sum.cell(row=row_num, column=1, value="Grand Total")
    c2 = ws_sum.cell(row=row_num, column=2, value=tot_fwd_non)
    c3 = ws_sum.cell(row=row_num, column=3, value=tot_fwd_adh)
    c4 = ws_sum.cell(row=row_num, column=4, value=tot_fwd_total)
    c5 = ws_sum.cell(row=row_num, column=5, value=tot_fwd_pct)

    c7 = ws_sum.cell(row=row_num, column=7, value="Grand Total")
    c8 = ws_sum.cell(row=row_num, column=8, value=tot_rev_non)
    c9 = ws_sum.cell(row=row_num, column=9, value=tot_rev_adh)
    c10 = ws_sum.cell(row=row_num, column=10, value=tot_rev_total)
    c11 = ws_sum.cell(row=row_num, column=11, value=tot_rev_pct)

    for cell in [c1, c2, c3, c4, c5, c7, c8, c9, c10, c11]:
        cell.font = font_total
        cell.border = border_thin

    c1.alignment = Alignment(horizontal="center", vertical="center")
    c2.alignment = Alignment(horizontal="right", vertical="center")
    c3.alignment = Alignment(horizontal="right", vertical="center")
    c4.alignment = Alignment(horizontal="right", vertical="center")
    c5.alignment = Alignment(horizontal="right", vertical="center")

    c7.alignment = Alignment(horizontal="center", vertical="center")
    c8.alignment = Alignment(horizontal="right", vertical="center")
    c9.alignment = Alignment(horizontal="right", vertical="center")
    c10.alignment = Alignment(horizontal="right", vertical="center")
    c11.alignment = Alignment(horizontal="right", vertical="center")

    c2.number_format = "#,##0"
    c3.number_format = "#,##0"
    c4.number_format = "#,##0"
    c5.number_format = "0.0%"

    c8.number_format = "#,##0"
    c9.number_format = "#,##0"
    c10.number_format = "#,##0"
    c11.number_format = "0.0%"

    f_fill, f_font = get_adherence_color_styles(tot_fwd_pct)
    c5.fill = f_fill
    c5.font = f_font

    r_fill, r_font = get_adherence_color_styles(tot_rev_pct)
    c11.fill = r_fill
    c11.font = r_font

    # Set column widths
    ws_sum.column_dimensions['F'].width = 3
    for col in ws_sum.columns:
        col_letter = get_column_letter(col[0].column)
        if col_letter != 'F':
            max_len = max(len(str(cell.value or '')) for cell in col)
            ws_sum.column_dimensions[col_letter].width = max(max_len + 3, 12)

    # --- Tab 2: Raw ---
    ws_raw = out_wb.create_sheet(title="Raw")
    ws_raw.sheet_view.showGridLines = True

    if raw_headers:
        for c_idx, h in enumerate(raw_headers, start=1):
            cell = ws_raw.cell(row=1, column=c_idx, value=h)
            cell.font = Font(name="Calibri", size=10, bold=True, color="FFFFFF")
            cell.fill = PatternFill(start_color="1F497D", end_color="1F497D", fill_type="solid")
            cell.alignment = Alignment(horizontal="center", vertical="center")

    for r_idx, row_vals in enumerate(filtered_raw_records, start=2):
        for c_idx, val in enumerate(row_vals, start=1):
            ws_raw.cell(row=r_idx, column=c_idx, value=val)

    for col in ws_raw.columns:
        col_letter = get_column_letter(col[0].column)
        max_len = 0
        for cell in col:
            val_str = str(cell.value or '')
            max_len = max(max_len, len(val_str))
        ws_raw.column_dimensions[col_letter].width = min(max(max_len + 3, 10), 35)

    out_wb.save(output_path)
    log.info(f"Successfully generated Second Attempt Adherence report: {output_path.name}")
    return str(output_path)
