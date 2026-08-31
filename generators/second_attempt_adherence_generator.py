"""
Second Attempt Adherence Report Generator
=========================================
Filters 2nd Attempt Adherence data for target DCs using dc_config.

Generates a 2-tab Excel output:
  1. 'Summary': Compact side-by-side FWD & REV DC metrics with short headers and narrow column widths.
     Conditional color fills for Adh %:
       - < 85%: Red
       - 85% - 95%: Yellow
       - > 95%: Green
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

try:
    from config.dc_config import ALLOWED_DCS_SET
    TARGET_DCS = set(dc.upper() for dc in ALLOWED_DCS_SET if dc.upper() != 'ALL')
except ImportError:
    try:
        from dc_config import ALLOWED_DCS_SET
        TARGET_DCS = set(dc.upper() for dc in ALLOWED_DCS_SET if dc.upper() != 'ALL')
    except ImportError:
        TARGET_DCS = {'ALG', 'AYP', 'DEO', 'JHS', 'JNP', 'KNP', 'MAU', 'MRZ', 'MTH', 'MZN', 'RBR', 'SPR', 'VNS'}

try:
    from python_calamine import CalamineWorkbook
except ImportError:
    CalamineWorkbook = None

log = logging.getLogger("ei_stream_server.2nd_attempt_adherence")
logging.basicConfig(level=logging.INFO, format="%(asctime)s [%(levelname)s] %(message)s")


def safe_int(val: Any) -> int:
    """Safely convert any value (including NaN, None, empty string) to int without crashing."""
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
    """Safely convert any value (including NaN, None, empty string) to float without crashing."""
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
    """
    Conditional styling for Adherence %:
      - Below 85% (< 0.85): Soft Red fill with dark red text
      - 85% to 95% (0.85 <= pct <= 0.95): Soft Yellow fill with dark amber text
      - Above 95% (> 0.95): Soft Green fill with dark green text
    """
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


def generate_second_attempt_adherence_report(input_path: Union[str, Path], output_path: Union[str, Path]) -> Path:
    input_path = Path(input_path)
    output_path = Path(output_path)
    log.info(f"Processing Second Attempt Adherence report for: {input_path.name}")

    if not input_path.exists():
        raise FileNotFoundError(f"Input file not found: {input_path}")

    # 1. Read Sheets (Calamine or pandas fallback)
    fwd_rows = []
    rev_rows = []
    summary_rows = []

    if CalamineWorkbook:
        try:
            wb_cal = CalamineWorkbook.from_path(str(input_path))
            sheet_map = {s.lower(): s for s in wb_cal.sheet_names}
            
            if "summary" in sheet_map:
                summary_rows = wb_cal.get_sheet_by_name(sheet_map["summary"]).to_python()
            if "fwd" in sheet_map:
                fwd_rows = wb_cal.get_sheet_by_name(sheet_map["fwd"]).to_python()
            if "rev" in sheet_map:
                rev_rows = wb_cal.get_sheet_by_name(sheet_map["rev"]).to_python()
        except Exception:
            wb_cal = None

    if not summary_rows and not fwd_rows and not rev_rows:
        import pandas as pd
        engine = 'pyxlsb' if input_path.suffix.lower() == '.xlsb' else None
        try:
            xls = pd.ExcelFile(str(input_path), engine=engine) if engine else pd.ExcelFile(str(input_path))
        except Exception:
            xls = pd.ExcelFile(str(input_path))
        sheet_map = {s.lower(): s for s in xls.sheet_names}
        if "summary" in sheet_map:
            summary_rows = pd.read_excel(xls, sheet_map["summary"], header=None).values.tolist()
        if "fwd" in sheet_map:
            fwd_df = pd.read_excel(xls, sheet_map["fwd"])
            fwd_rows = [fwd_df.columns.tolist()] + fwd_df.values.tolist()
        if "rev" in sheet_map:
            rev_df = pd.read_excel(xls, sheet_map["rev"])
            rev_rows = [rev_df.columns.tolist()] + rev_df.values.tolist()

    # 2. Parse & Filter Summary Data Side-by-Side
    fwd_dict = {}
    rev_dict = {}

    for r in summary_rows[2:]:
        # FWD (Cols 0-4)
        if len(r) >= 5 and r[0] is not None:
            dc = str(r[0]).strip().upper()
            if dc in TARGET_DCS:
                non_adh = safe_int(r[1])
                adh = safe_int(r[2])
                total = safe_int(r[3]) if r[3] is not None else (non_adh + adh)
                pct = safe_float(r[4]) if (r[4] is not None and not math.isnan(safe_float(r[4]))) else (adh / total if total > 0 else 0.0)
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
                total = safe_int(r[9]) if r[9] is not None else (non_adh + adh)
                pct = safe_float(r[10]) if (r[10] is not None and not math.isnan(safe_float(r[10]))) else (adh / total if total > 0 else 0.0)
                rev_dict[dc] = {
                    "non_adherence": non_adh,
                    "adherence": adh,
                    "grand_total": total,
                    "adherence_pct": pct
                }

    sorted_dcs = sorted(list(set(fwd_dict.keys()) | set(rev_dict.keys())))

    # 3. Filter Raw Data (FWD & REV) by Source_DC
    filtered_raw_records = []
    raw_headers = []

    def _get_dc_idx(h_row):
        lower = [str(h).strip().lower() for h in h_row if h is not None]
        for candidate in ['source_dc', 'source dc', 'dc', 'sourcedc', 'hub']:
            if candidate in lower:
                return lower.index(candidate)
        return 0

    # Process FWD Raw
    if fwd_rows:
        headers = [str(h).strip() if h is not None else "" for h in fwd_rows[0]]
        raw_headers = ["Flow_Type"] + headers
        dc_idx = _get_dc_idx(headers)
        for row in fwd_rows[1:]:
            if len(row) > dc_idx and row[dc_idx] is not None:
                dc_val = str(row[dc_idx]).strip().upper()
                if dc_val in TARGET_DCS:
                    filtered_raw_records.append(["FWD"] + list(row))

    # Process REV Raw
    if rev_rows:
        headers = [str(h).strip() if h is not None else "" for h in rev_rows[0]]
        if not raw_headers:
            raw_headers = ["Flow_Type"] + headers
        dc_idx = _get_dc_idx(headers)
        for row in rev_rows[1:]:
            if len(row) > dc_idx and row[dc_idx] is not None:
                dc_val = str(row[dc_idx]).strip().upper()
                if dc_val in TARGET_DCS:
                    filtered_raw_records.append(["REV"] + list(row))

    log.info(f"Summary Dotted DC Items: FWD={len(fwd_dict)}, REV={len(rev_dict)}")
    log.info(f"Total Filtered Raw Records: {len(filtered_raw_records)}")

    # 4. Build Output Workbook using openpyxl
    out_wb = openpyxl.Workbook()
    
    # --- Tab 1: Summary ---
    ws_sum = out_wb.active
    ws_sum.title = "Summary"
    ws_sum.views.sheetView[0].showGridLines = True

    # Styling definitions
    font_top_header = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    font_fwd_sub = Font(name="Calibri", size=9, bold=True, color="1F497D")
    font_rev_sub = Font(name="Calibri", size=9, bold=True, color="375623")
    font_regular = Font(name="Calibri", size=9)

    # Top Header Fills
    fill_fwd_top = PatternFill(start_color="2F5597", end_color="2F5597", fill_type="solid")
    fill_rev_top = PatternFill(start_color="375623", end_color="375623", fill_type="solid")

    # Sub-header Fills
    fill_fwd_sub = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")
    fill_rev_sub = PatternFill(start_color="E2EFDA", end_color="E2EFDA", fill_type="solid")

    border_thin = Border(
        left=Side(style="thin", color="D9D9D9"),
        right=Side(style="thin", color="D9D9D9"),
        top=Side(style="thin", color="D9D9D9"),
        bottom=Side(style="thin", color="D9D9D9")
    )

    # Row 1: Merged Top Headers ("FWD" on A1:E1, "REV" on G1:K1)
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

    # Row 2: Short Sub-headers for FWD (Cols 1-5) and REV (Cols 7-11)
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

    # Rows 3+: Side-by-Side Data Rows (No Total Row)
    current_row = 3
    for dc in sorted_dcs:
        fwd_item = fwd_dict.get(dc, {"non_adherence": 0, "adherence": 0, "grand_total": 0, "adherence_pct": 0.0})
        rev_item = rev_dict.get(dc, {"non_adherence": 0, "adherence": 0, "grand_total": 0, "adherence_pct": 0.0})

        # FWD Data (Cols 1-5)
        ws_sum.cell(row=current_row, column=1, value=dc).font = font_regular
        ws_sum.cell(row=current_row, column=2, value=fwd_item["non_adherence"]).number_format = "#,##0"
        ws_sum.cell(row=current_row, column=3, value=fwd_item["adherence"]).number_format = "#,##0"
        ws_sum.cell(row=current_row, column=4, value=fwd_item["grand_total"]).number_format = "#,##0"
        pct_fwd = ws_sum.cell(row=current_row, column=5, value=fwd_item["adherence_pct"])
        pct_fwd.number_format = "0.0%"
        fwd_fill, fwd_font = get_adherence_color_styles(fwd_item["adherence_pct"])
        pct_fwd.fill = fwd_fill
        pct_fwd.font = fwd_font

        # REV Data (Cols 7-11)
        ws_sum.cell(row=current_row, column=7, value=dc).font = font_regular
        ws_sum.cell(row=current_row, column=8, value=rev_item["non_adherence"]).number_format = "#,##0"
        ws_sum.cell(row=current_row, column=9, value=rev_item["adherence"]).number_format = "#,##0"
        ws_sum.cell(row=current_row, column=10, value=rev_item["grand_total"]).number_format = "#,##0"
        pct_rev = ws_sum.cell(row=current_row, column=11, value=rev_item["adherence_pct"])
        pct_rev.number_format = "0.0%"
        rev_fill, rev_font = get_adherence_color_styles(rev_item["adherence_pct"])
        pct_rev.fill = rev_fill
        pct_rev.font = rev_font

        # Borders and Alignment
        for c in range(1, 6):
            cell = ws_sum.cell(row=current_row, column=c)
            cell.border = border_thin
            if c > 1: cell.alignment = Alignment(horizontal="right")

        for c in range(7, 12):
            cell = ws_sum.cell(row=current_row, column=c)
            cell.border = border_thin
            if c > 1: cell.alignment = Alignment(horizontal="right")

        current_row += 1

    # Compact Column Widths for Summary Tab
    compact_widths = {
        "A": 9,  "B": 10, "C": 10, "D": 9,  "E": 9,
        "F": 3,
        "G": 9,  "H": 10, "I": 10, "J": 9,  "K": 9
    }
    for col_letter, width in compact_widths.items():
        ws_sum.column_dimensions[col_letter].width = width

    # --- Tab 2: Raw ---
    ws_raw = out_wb.create_sheet(title="Raw")
    ws_raw.views.sheetView[0].showGridLines = True

    if raw_headers:
        ws_raw.append(raw_headers)
        font_raw_header = Font(name="Calibri", size=10, bold=True, color="FFFFFF")
        fill_raw_header = PatternFill(start_color="334155", end_color="334155", fill_type="solid")
        for col_num in range(1, len(raw_headers) + 1):
            cell = ws_raw.cell(row=1, column=col_num)
            cell.font = font_raw_header
            cell.fill = fill_raw_header
            cell.alignment = Alignment(horizontal="center", vertical="center")

            # Fast header width estimation
            col_letter = get_column_letter(col_num)
            h_len = len(str(raw_headers[col_num - 1]))
            ws_raw.column_dimensions[col_letter].width = min(max(h_len + 4, 12), 35)

    for record in filtered_raw_records:
        ws_raw.append(record)

    # Save Output
    output_path.parent.mkdir(parents=True, exist_ok=True)
    out_wb.save(output_path)
    try:
        out_wb.close()
    except Exception:
        pass
    log.info(f"Successfully generated 2nd Attempt Adherence report: {output_path} ({output_path.stat().st_size} bytes)")
    return output_path


if __name__ == "__main__":
    import argparse
    parser = argparse.ArgumentParser(description="Generate 2nd Attempt Adherence Report")
    parser.add_argument("input_file", nargs="?", default=r"C:\Users\User\Desktop\server\2nd attempt adherence report\Second Attempt Adherence_01-Aug-2026(FWD&REV).xlsb")
    parser.add_argument("output_file", nargs="?", default=r"C:\Users\User\Desktop\server\2nd attempt adherence report\Output_2nd_Attempt_Adherence.xlsx")
    args = parser.parse_args()

    out = generate_second_attempt_adherence_report(args.input_file, args.output_file)
    print(f"\nReport Generated Successfully! Saved to: {out}")
