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
import tempfile
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
    import gc
    input_path = Path(input_path)
    output_path = Path(output_path)
    log.info(f"Processing Second Attempt Adherence report (streaming mode) for: {input_path.name}")

    if not input_path.exists():
        raise FileNotFoundError(f"Input file not found: {input_path}")

    # 1. Parse Summary Data Side-by-Side
    fwd_dict = {}
    rev_dict = {}

    summary_rows = []
    wb_cal = None
    if CalamineWorkbook:
        try:
            wb_cal = CalamineWorkbook.from_path(str(input_path))
            sheet_map = {s.lower(): s for s in wb_cal.sheet_names}
            if "summary" in sheet_map:
                summary_rows = wb_cal.get_sheet_by_name(sheet_map["summary"]).to_python()
        except Exception as cal_err:
            log.warning(f"Calamine summary read: {cal_err}")
            wb_cal = None

    if not summary_rows:
        try:
            in_wb = openpyxl.load_workbook(str(input_path), read_only=True, data_only=True)
            sheet_map = {s.lower(): s for s in in_wb.sheetnames}
            if "summary" in sheet_map:
                ws_s = in_wb[sheet_map["summary"]]
                summary_rows = [list(r) for r in ws_s.iter_rows(values_only=True)]
            in_wb.close()
        except Exception as ox_err:
            log.warning(f"openpyxl summary read: {ox_err}")

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

    # 2. Build Output Workbook
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

    # Row 1: Merged Top Headers
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

    # Row 2: Sub-headers
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

    # Rows 3+: Data Rows
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

    # 3. Stream Raw Data Directly to Zip XML
    import io
    import zipfile

    temp_sum = io.BytesIO()
    out_wb.save(temp_sum)
    temp_sum.seek(0)
    try:
        out_wb.close()
    except Exception:
        pass
    del out_wb
    gc.collect()

    z_in = zipfile.ZipFile(temp_sum, "r")
    z_out = zipfile.ZipFile(output_path, "w", compression=zipfile.ZIP_DEFLATED)

    for item in z_in.infolist():
        if item.filename == "[Content_Types].xml":
            ct = z_in.read(item.filename).decode("utf-8")
            ct = ct.replace("</Types>", '<Override PartName="/xl/worksheets/sheet2.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/></Types>')
            z_out.writestr(item.filename, ct)
        elif item.filename == "xl/workbook.xml":
            wb_xml = z_in.read(item.filename).decode("utf-8")
            wb_xml = wb_xml.replace("</sheets>", '<sheet name="Raw" sheetId="2" r:id="rId2"/></sheets>')
            z_out.writestr(item.filename, wb_xml)
        elif item.filename == "xl/_rels/workbook.xml.rels":
            wb_rels = z_in.read(item.filename).decode("utf-8")
            wb_rels = wb_rels.replace("</Relationships>", '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet2.xml"/></Relationships>')
            z_out.writestr(item.filename, wb_rels)
        else:
            z_out.writestr(item, z_in.read(item.filename))

    def _get_dc_idx(h_row):
        lower = [str(h).strip().lower() for h in h_row if h is not None]
        for candidate in ["source_dc", "source dc", "dc", "sourcedc", "hub"]:
            if candidate in lower:
                return lower.index(candidate)
        return 0

    def esc(val):
        if val is None:
            return ""
        s = str(val)
        return s.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;").replace('"', "&quot;")

    raw_headers = []
    fwd_rows_iter = []
    rev_rows_iter = []
    raw_rows_iter = []
    dc_idx = 0

    if wb_cal:
        sheet_map = {s.lower(): s for s in wb_cal.sheet_names}
        if "fwd" in sheet_map:
            fwd_sheet = wb_cal.get_sheet_by_name(sheet_map["fwd"])
            fwd_rows_iter = fwd_sheet.iter_rows()
            h_fwd = next(fwd_rows_iter, None)
            if h_fwd:
                raw_headers = ["Flow_Type"] + [str(h).strip() if h is not None else "" for h in h_fwd]
                dc_idx = _get_dc_idx(h_fwd)
        if "rev" in sheet_map:
            rev_sheet = wb_cal.get_sheet_by_name(sheet_map["rev"])
            rev_rows_iter = rev_sheet.iter_rows()
            h_rev = next(rev_rows_iter, None)
            if not raw_headers and h_rev:
                raw_headers = ["Flow_Type"] + [str(h).strip() if h is not None else "" for h in h_rev]
                dc_idx = _get_dc_idx(h_rev)
        if not fwd_rows_iter and not rev_rows_iter:
            for cand in ["raw", "raw_data", "data", "sheet1"]:
                if cand in sheet_map:
                    raw_sheet = wb_cal.get_sheet_by_name(sheet_map[cand])
                    raw_rows_iter = raw_sheet.iter_rows()
                    h_raw = next(raw_rows_iter, None)
                    if h_raw:
                        raw_headers = [str(h).strip() if h is not None else "" for h in h_raw]
                        dc_idx = _get_dc_idx(h_raw)
                    break
    else:
        in_wb = openpyxl.load_workbook(str(input_path), read_only=True, data_only=True)
        sheet_map = {s.lower(): s for s in in_wb.sheetnames}
        if "fwd" in sheet_map:
            ws_fwd = in_wb[sheet_map["fwd"]]
            fwd_rows_iter = ws_fwd.iter_rows(values_only=True)
            h_fwd = next(fwd_rows_iter, None)
            if h_fwd:
                raw_headers = ["Flow_Type"] + [str(h).strip() if h is not None else "" for h in h_fwd]
                dc_idx = _get_dc_idx(h_fwd)
        if "rev" in sheet_map:
            ws_rev = in_wb[sheet_map["rev"]]
            rev_rows_iter = ws_rev.iter_rows(values_only=True)
            h_rev = next(rev_rows_iter, None)
            if not raw_headers and h_rev:
                raw_headers = ["Flow_Type"] + [str(h).strip() if h is not None else "" for h in h_rev]
                dc_idx = _get_dc_idx(h_rev)
        if not fwd_rows_iter and not rev_rows_iter:
            for cand in ["raw", "raw_data", "data", "sheet1"]:
                if cand in sheet_map:
                    ws_raw = in_wb[sheet_map[cand]]
                    raw_rows_iter = ws_raw.iter_rows(values_only=True)
                    h_raw = next(raw_rows_iter, None)
                    if h_raw:
                        raw_headers = [str(h).strip() if h is not None else "" for h in h_raw]
                        dc_idx = _get_dc_idx(h_raw)
                    break

    raw_record_count = 0
    temp_raw_xml = Path(tempfile.gettempdir()) / f"temp_raw_2nd_{output_path.stem}.xml"
    with open(temp_raw_xml, "wb") as f:
        f.write(b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData>')
        row_num = 1

        # Header
        if raw_headers:
            hdr_xml = ['<row r="1">']
            for c_i, h in enumerate(raw_headers, 1):
                col_let = get_column_letter(c_i)
                hdr_xml.append(f'<c r="{col_let}1" t="inlineStr"><is><t>{esc(h)}</t></is></c>')
            hdr_xml.append("</row>")
            f.write("".join(hdr_xml).encode("utf-8"))
            row_num += 1

        chunk = []

        # FWD
        for row in fwd_rows_iter:
            if len(row) > dc_idx and row[dc_idx] is not None:
                if str(row[dc_idx]).strip().upper() in TARGET_DCS:
                    r_xml = [f'<row r="{row_num}"><c r="A{row_num}" t="inlineStr"><is><t>FWD</t></is></c>']
                    for c_i, val in enumerate(row, 2):
                        col_let = get_column_letter(c_i)
                        if isinstance(val, (int, float)) and not math.isnan(val) and not math.isinf(val):
                            r_xml.append(f'<c r="{col_let}{row_num}"><v>{val}</v></c>')
                        else:
                            r_xml.append(f'<c r="{col_let}{row_num}" t="inlineStr"><is><t>{esc(val)}</t></is></c>')
                    r_xml.append("</row>")
                    chunk.append("".join(r_xml))
                    row_num += 1
                    raw_record_count += 1
                    if len(chunk) >= 1000:
                        f.write("".join(chunk).encode("utf-8"))
                        chunk.clear()

        # REV
        for row in rev_rows_iter:
            if len(row) > dc_idx and row[dc_idx] is not None:
                if str(row[dc_idx]).strip().upper() in TARGET_DCS:
                    r_xml = [f'<row r="{row_num}"><c r="A{row_num}" t="inlineStr"><is><t>REV</t></is></c>']
                    for c_i, val in enumerate(row, 2):
                        col_let = get_column_letter(c_i)
                        if isinstance(val, (int, float)) and not math.isnan(val) and not math.isinf(val):
                            r_xml.append(f'<c r="{col_let}{row_num}"><v>{val}</v></c>')
                        else:
                            r_xml.append(f'<c r="{col_let}{row_num}" t="inlineStr"><is><t>{esc(val)}</t></is></c>')
                    r_xml.append("</row>")
                    chunk.append("".join(r_xml))
                    row_num += 1
                    raw_record_count += 1
                    if len(chunk) >= 1000:
                        f.write("".join(chunk).encode("utf-8"))
                        chunk.clear()

        # Raw (Combined)
        for row in raw_rows_iter:
            if len(row) > dc_idx and row[dc_idx] is not None:
                if str(row[dc_idx]).strip().upper() in TARGET_DCS:
                    r_xml = [f'<row r="{row_num}">']
                    for c_i, val in enumerate(row, 1):
                        col_let = get_column_letter(c_i)
                        if isinstance(val, (int, float)) and not math.isnan(val) and not math.isinf(val):
                            r_xml.append(f'<c r="{col_let}{row_num}"><v>{val}</v></c>')
                        else:
                            r_xml.append(f'<c r="{col_let}{row_num}" t="inlineStr"><is><t>{esc(val)}</t></is></c>')
                    r_xml.append("</row>")
                    chunk.append("".join(r_xml))
                    row_num += 1
                    raw_record_count += 1
                    if len(chunk) >= 1000:
                        f.write("".join(chunk).encode("utf-8"))
                        chunk.clear()

        if chunk:
            f.write("".join(chunk).encode("utf-8"))
            chunk.clear()

        f.write(b"</sheetData></worksheet>")

    with open(temp_raw_xml, "rb") as f_raw:
        z_out.writestr("xl/worksheets/sheet2.xml", f_raw.read())

    z_out.close()
    z_in.close()

    if temp_raw_xml.exists():
        try:
            temp_raw_xml.unlink()
        except Exception:
            pass

    log.info(f"Summary DC Items: FWD={len(fwd_dict)}, REV={len(rev_dict)}")
    log.info(f"Total Streamed Raw Records: {raw_record_count}")
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
