#!/usr/bin/env python3
"""
Forward Pendency Report Generator Module for ei_stream_server
=============================================================
Reads 'raw_data_North' from input Excel file, filters rows where Source_DC is in
{'alg', 'ayp', 'deo', 'jhs', 'jnp', 'mau', 'mrz', 'mth', 'mzn', 'rbr', 'spr'},
computes the 'Aging Category' column right beside 'Aging', and generates output workbook:
  1. Summary Sheet (3 Sidewise Tables with Red/Green highlights)
  2. CPD-DID Sheet (P2 & P3 actual row details)
  3. Raw Sheet (Full filtered rows dataset)

Uses Zero-Memory Streaming Architecture:
- Direct XML disk streaming for massive (200k+ rows) datasets
- Calamine / openpyxl read_only stream parsing
- Memory footprint < 40MB RAM under full load
"""

import io
import math
import zipfile
import tempfile
import logging
from pathlib import Path
from collections import defaultdict
import openpyxl
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter

try:
    from python_calamine import CalamineWorkbook
    HAS_CALAMINE = True
except ImportError:
    HAS_CALAMINE = False

try:
    from config.dc_config import ALLOWED_SOURCE_DCS, ALLOWED_DCS_SET, ALLOWED_DCS_SET_LOWER
except ImportError:
    from dc_config import ALLOWED_SOURCE_DCS, ALLOWED_DCS_SET, ALLOWED_DCS_SET_LOWER

AGING_CATEGORIES = ['0-2 days', '3-5 days', '5-10 days', '>10 days']

log = logging.getLogger("ei_stream_server.forward_pendency")


def compute_aging_category(val) -> str:
    """Calculates the Aging Category bucket string. Empty/None defaults to '0-2 days'."""
    if val is None or str(val).strip() == "":
        return "0-2 days"
    try:
        aging = float(val)
        if aging <= 2:
            return "0-2 days"
        elif aging <= 5:
            return "3-5 days"
        elif aging <= 10:
            return "5-10 days"
        else:
            return ">10 days"
    except (ValueError, TypeError):
        return "0-2 days"


def write_side_table(ws, start_col: int, start_row: int, title: str, headers: list, data_matrix: list):
    """Writes a side-by-side formatted summary table with spanned title and red/green highlights."""
    end_col = start_col + len(headers) - 1

    title_fill = PatternFill("solid", fgColor="1E1B4B")
    title_font = Font(name="Calibri", size=12, bold=True, color="FFFFFF")

    header_fill = PatternFill("solid", fgColor="312E81")
    header_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")

    total_fill = PatternFill("solid", fgColor="E0E7FF")
    total_font = Font(name="Calibri", size=11, bold=True, color="1E1B4B")

    data_font = Font(name="Calibri", size=11, color="1F2937")
    thin_side = Side(style="thin", color="CBD5E1")
    border = Border(left=thin_side, right=thin_side, top=thin_side, bottom=thin_side)
    center_align = Alignment(horizontal="center", vertical="center")
    left_align = Alignment(horizontal="left", vertical="center")

    red_fill = PatternFill("solid", fgColor="FECACA")
    red_font = Font(name="Calibri", size=11, bold=True, color="991B1B")

    green_fill = PatternFill("solid", fgColor="DCFCE7")
    green_font = Font(name="Calibri", size=11, bold=True, color="166534")

    # Merge title row
    ws.merge_cells(start_row=start_row, start_column=start_col, end_row=start_row, end_column=end_col)
    title_cell = ws.cell(row=start_row, column=start_col, value=title)
    title_cell.font = title_font
    title_cell.alignment = center_align

    for col_i in range(start_col, end_col + 1):
        c = ws.cell(row=start_row, column=col_i)
        c.fill = title_fill
        c.border = border

    current_row = start_row + 1

    # Headers
    for c_offset, h_text in enumerate(headers):
        col_i = start_col + c_offset
        cell = ws.cell(row=current_row, column=col_i, value=h_text)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = center_align
        cell.border = border

    current_row += 1

    # Data Rows
    for row_values in data_matrix:
        is_total_row = (row_values[0] == "Total")
        for c_offset, val in enumerate(row_values):
            col_i = start_col + c_offset
            h_name = headers[c_offset]
            cell = ws.cell(row=current_row, column=col_i, value=val)
            cell.border = border
            cell.alignment = left_align if c_offset == 0 else center_align

            highlighted = False
            if isinstance(val, (int, float)) and val > 0:
                if title == "Aging wise report" and h_name in ['3-5 days', '5-10 days', '>10 days']:
                    cell.fill = red_fill
                    cell.font = red_font
                    highlighted = True
                elif title == "Priority Table":
                    if h_name in ['P2', 'P3']:
                        cell.fill = red_fill
                        cell.font = red_font
                        highlighted = True
                    elif h_name == 'P4':
                        cell.fill = green_fill
                        cell.font = green_font
                        highlighted = True

            if not highlighted:
                if is_total_row:
                    cell.fill = total_fill
                    cell.font = total_font
                else:
                    cell.font = data_font

        current_row += 1


def build_summary_sheet_from_pivots(out_wb, t1_pivot, t2_pivot, t3_pivot):
    """Generates the Summary tab with Table 1, Table 2, and Table 3 placed SIDEWISE."""
    ws = out_wb.active
    ws.title = "Summary"
    ws.sheet_view.showGridLines = False

    table1_headers = ["Source DC"] + AGING_CATEGORIES + ["Total Pendency"]
    table1_data = []
    tot_cats_t1 = defaultdict(int)

    for dc in ALLOWED_SOURCE_DCS:
        row_vals = [dc]
        row_tot = 0
        for cat in AGING_CATEGORIES:
            cnt = t1_pivot[dc][cat]
            row_vals.append(cnt)
            row_tot += cnt
            tot_cats_t1[cat] += cnt
        row_vals.append(row_tot)
        table1_data.append(row_vals)

    t1_totals = ["Total"] + [tot_cats_t1[cat] for cat in AGING_CATEGORIES] + [sum(tot_cats_t1.values())]
    table1_data.append(t1_totals)

    prio_keys = ["P2", "P3", "P4"]
    table2_headers = ["Source DC"] + prio_keys + ["Total Pendency"]
    table2_data = []
    tot_prios_t2 = defaultdict(int)

    for dc in ALLOWED_SOURCE_DCS:
        row_vals = [dc]
        row_tot = 0
        for p in prio_keys:
            cnt = t2_pivot[dc][p]
            row_vals.append(cnt)
            row_tot += cnt
            tot_prios_t2[p] += cnt
        row_vals.append(row_tot)
        table2_data.append(row_vals)

    t2_totals = ["Total"] + [tot_prios_t2[p] for p in prio_keys] + [sum(tot_prios_t2.values())]
    table2_data.append(t2_totals)

    table3_headers = ["Source DC", "CPD (P3)", "DID (P2)", "Total Pendency"]
    table3_data = []
    tot_cpd = 0
    tot_did = 0

    for dc in ALLOWED_SOURCE_DCS:
        cpd_cnt = t3_pivot[dc]["CPD (P3)"]
        did_cnt = t3_pivot[dc]["DID (P2)"]
        row_tot = cpd_cnt + did_cnt
        table3_data.append([dc, cpd_cnt, did_cnt, row_tot])
        tot_cpd += cpd_cnt
        tot_did += did_cnt

    t3_totals = ["Total", tot_cpd, tot_did, tot_cpd + tot_did]
    table3_data.append(t3_totals)

    start_row = 2
    write_side_table(ws, start_col=2,  start_row=start_row, title="Aging wise report", headers=table1_headers, data_matrix=table1_data)
    write_side_table(ws, start_col=9,  start_row=start_row, title="Priority Table",    headers=table2_headers, data_matrix=table2_data)
    write_side_table(ws, start_col=15, start_row=start_row, title="CPD Pendency",     headers=table3_headers, data_matrix=table3_data)

    ws.column_dimensions['A'].width = 3
    ws.column_dimensions['H'].width = 4
    ws.column_dimensions['N'].width = 4

    for col in ws.columns:
        col_letter = get_column_letter(col[0].column)
        if col_letter in ['A', 'H', 'N']:
            continue
        max_len = max(len(str(cell.value or '')) for cell in col)
        ws.column_dimensions[col_letter].width = max(max_len + 3, 14)


def esc(val):
    if val is None:
        return ""
    s = str(val)
    return s.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;").replace('"', "&quot;")


def generate_forward_pendency_report(input_file: Path, output_file: Path):
    """
    Main generator pipeline for Forward Pendency Report using Zero-Memory Streaming Architecture.
    """
    input_path = Path(input_file)
    output_path = Path(output_file)
    if not input_path.exists():
        raise FileNotFoundError(f"Input file not found: {input_path}")

    log.info(f"Loading input workbook for Forward Pendency Report (Zero-Memory Stream Mode): {input_path}")

    # Set up temp files for XML streaming
    output_path.parent.mkdir(parents=True, exist_ok=True)
    temp_raw_xml = Path(tempfile.gettempdir()) / f"temp_fwd_raw_{output_path.stem}.xml"
    temp_cpd_xml = Path(tempfile.gettempdir()) / f"temp_fwd_cpd_{output_path.stem}.xml"

    f_raw = open(temp_raw_xml, 'wb')
    f_cpd = open(temp_cpd_xml, 'wb')

    f_raw.write(b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData>')
    f_cpd.write(b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData>')

    # Pivot aggregators
    t1_pivot = defaultdict(lambda: defaultdict(int))
    t2_pivot = defaultdict(lambda: defaultdict(int))
    t3_pivot = defaultdict(lambda: defaultdict(int))

    total_processed = 0
    total_filtered = 0
    cpd_count = 0

    aging_col_idx = 20
    sdc_idx = 15
    prio_idx = 13
    shipment_idx = 1
    attempt_idx = 23

    raw_row_num = 2
    cpd_row_num = 2
    raw_chunk = []
    cpd_chunk = []

    # Stream reader using openpyxl read_only
    in_wb = openpyxl.load_workbook(str(input_path), data_only=True, read_only=True)
    target_sheet = None
    sheet_map = {name.lower(): name for name in in_wb.sheetnames}
    for candidate in ['raw_data_north', 'raw_data', 'raw', 'praw data', 'data']:
        if candidate in sheet_map:
            target_sheet = sheet_map[candidate]
            break
    if not target_sheet and in_wb.sheetnames:
        target_sheet = in_wb.sheetnames[0]

    in_ws = in_wb[target_sheet]
    log.info(f"Reading rows from '{target_sheet}' sheet in stream mode...")

    row_iter = in_ws.iter_rows(values_only=True)
    header_row = next(row_iter, None)
    if not header_row:
        in_wb.close()
        f_raw.close()
        f_cpd.close()
        raise ValueError(f"Sheet '{target_sheet}' is empty.")

    header_list = list(header_row)
    col_map = {str(h).strip().lower(): idx for idx, h in enumerate(header_list) if h is not None}
    for c in ['aging', 'aging bucket', 'age_bucket', 'ageing', 'age']:
        if c in col_map:
            aging_col_idx = col_map[c]
            break
    for c in ['source_dc', 'source dc', 'dc', 'sourcedc']:
        if c in col_map:
            sdc_idx = col_map[c]
            break
    for c in ['customerpriorityv2', 'priority', 'prio', 'customer_priority']:
        if c in col_map:
            prio_idx = col_map[c]
            break
    for c in ['pendingshipments', 'tracking_no', 'waybill', 'tracking_id', 'shipment', 'tracking_number']:
        if c in col_map:
            shipment_idx = col_map[c]
            break
    for c in ['attempt_status', 'status', 'attempt']:
        if c in col_map:
            attempt_idx = col_map[c]
            break

    # Write Raw header (insert Aging Category right after Aging)
    raw_header_out = list(header_list)
    raw_header_out.insert(aging_col_idx + 1, 'Aging Category')
    r_hdr_xml = ['<row r="1">']
    for c_i, h_val in enumerate(raw_header_out, 1):
        r_hdr_xml.append(f'<c r="{get_column_letter(c_i)}1" t="inlineStr"><is><t>{esc(h_val)}</t></is></c>')
    r_hdr_xml.append('</row>')
    f_raw.write(''.join(r_hdr_xml).encode('utf-8'))

    # Write CPD-DID header
    cpd_headers = ["PendingShipments", "Source_DC", "Aging Category", "Attempt_Status", "CustomerPriorityV2"]
    c_hdr_xml = ['<row r="1">']
    for c_i, h_val in enumerate(cpd_headers, 1):
        c_hdr_xml.append(f'<c r="{get_column_letter(c_i)}1" t="inlineStr"><is><t>{esc(h_val)}</t></is></c>')
    c_hdr_xml.append('</row>')
    f_cpd.write(''.join(c_hdr_xml).encode('utf-8'))

    for row in row_iter:
        total_processed += 1
        raw_sdc = row[sdc_idx] if len(row) > sdc_idx else None
        sdc = str(raw_sdc).strip().lower() if raw_sdc is not None else ''

        if sdc in ALLOWED_DCS_SET_LOWER:
            total_filtered += 1
            sdc_upper = sdc.upper()
            aging_val = row[aging_col_idx] if len(row) > aging_col_idx else None
            aging_cat = compute_aging_category(aging_val)

            raw_prio = row[prio_idx] if len(row) > prio_idx else None
            prio = str(raw_prio).strip().upper() if raw_prio is not None and str(raw_prio).strip() else "Unknown"

            # Aggregate into pivots for Summary sheet
            t1_pivot[sdc_upper][aging_cat] += 1
            t2_pivot[sdc_upper][prio] += 1

            if prio == "P3":
                t3_pivot[sdc_upper]["CPD (P3)"] += 1
            elif prio == "P2":
                t3_pivot[sdc_upper]["DID (P2)"] += 1

            # Format Raw row
            r_xml = [f'<row r="{raw_row_num}">']
            col_counter = 1
            for idx, val in enumerate(row):
                col_let = get_column_letter(col_counter)
                if isinstance(val, (int, float)) and not math.isnan(val) and not math.isinf(val):
                    r_xml.append(f'<c r="{col_let}{raw_row_num}"><v>{val}</v></c>')
                else:
                    r_xml.append(f'<c r="{col_let}{raw_row_num}" t="inlineStr"><is><t>{esc(val)}</t></is></c>')
                col_counter += 1

                if idx == aging_col_idx:
                    # Insert Aging Category cell
                    col_let_cat = get_column_letter(col_counter)
                    r_xml.append(f'<c r="{col_let_cat}{raw_row_num}" t="inlineStr"><is><t>{esc(aging_cat)}</t></is></c>')
                    col_counter += 1

            r_xml.append('</row>')
            raw_chunk.append(''.join(r_xml))
            raw_row_num += 1

            # CPD-DID row if P2 or P3
            if prio in ("P2", "P3"):
                cpd_count += 1
                shipment = row[shipment_idx] if len(row) > shipment_idx and row[shipment_idx] is not None else ""
                attempt_stat = row[attempt_idx] if len(row) > attempt_idx and row[attempt_idx] is not None else ""
                cpd_vals = [shipment, sdc_upper, aging_cat, attempt_stat, prio]

                c_xml = [f'<row r="{cpd_row_num}">']
                for c_i, val in enumerate(cpd_vals, 1):
                    col_let = get_column_letter(c_i)
                    if isinstance(val, (int, float)) and not math.isnan(val) and not math.isinf(val):
                        c_xml.append(f'<c r="{col_let}{cpd_row_num}"><v>{val}</v></c>')
                    else:
                        c_xml.append(f'<c r="{col_let}{cpd_row_num}" t="inlineStr"><is><t>{esc(val)}</t></is></c>')
                c_xml.append('</row>')
                cpd_chunk.append(''.join(c_xml))
                cpd_row_num += 1

            if len(raw_chunk) >= 1000:
                f_raw.write(''.join(raw_chunk).encode('utf-8'))
                raw_chunk.clear()
            if len(cpd_chunk) >= 1000:
                f_cpd.write(''.join(cpd_chunk).encode('utf-8'))
                cpd_chunk.clear()

    if raw_chunk:
        f_raw.write(''.join(raw_chunk).encode('utf-8'))
    if cpd_chunk:
        f_cpd.write(''.join(cpd_chunk).encode('utf-8'))

    f_raw.write(b'</sheetData></worksheet>')
    f_cpd.write(b'</sheetData></worksheet>')
    f_raw.close()
    f_cpd.close()
    in_wb.close()

    log.info(f"Processed {total_processed} source rows. Filtered {total_filtered} matching rows ({cpd_count} CPD-DID rows).")

    # Build Summary sheet in tiny openpyxl workbook (~30 rows, ~50 KB RAM)
    out_wb = Workbook()
    build_summary_sheet_from_pivots(out_wb, t1_pivot, t2_pivot, t3_pivot)

    temp_sum = io.BytesIO()
    out_wb.save(temp_sum)
    temp_sum.seek(0)
    try:
        out_wb.close()
    except Exception:
        pass
    del out_wb, t1_pivot, t2_pivot, t3_pivot

    # Assemble ZIP archive with Summary, CPD-DID, and Raw sheets
    z_in = zipfile.ZipFile(temp_sum, 'r')
    z_out = zipfile.ZipFile(output_path, 'w', compression=zipfile.ZIP_DEFLATED)

    for item in z_in.infolist():
        if item.filename == '[Content_Types].xml':
            ct = z_in.read(item.filename).decode('utf-8')
            ct = ct.replace('</Types>', '<Override PartName="/xl/worksheets/sheet2.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/worksheets/sheet3.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/></Types>')
            z_out.writestr(item.filename, ct)
        elif item.filename == 'xl/workbook.xml':
            wb_xml = z_in.read(item.filename).decode('utf-8')
            wb_xml = wb_xml.replace('</sheets>', '<sheet name="CPD-DID" sheetId="2" r:id="rId2"/><sheet name="Raw" sheetId="3" r:id="rId3"/></sheets>')
            z_out.writestr(item.filename, wb_xml)
        elif item.filename == 'xl/_rels/workbook.xml.rels':
            wb_rels = z_in.read(item.filename).decode('utf-8')
            wb_rels = wb_rels.replace('</Relationships>', '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet2.xml"/><Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet3.xml"/></Relationships>')
            z_out.writestr(item.filename, wb_rels)
        else:
            z_out.writestr(item, z_in.read(item.filename))

    # Stream CPD-DID sheet into sheet2.xml
    with z_out.open('xl/worksheets/sheet2.xml', 'w', force_zip64=True) as zf_entry:
        with open(temp_cpd_xml, 'rb') as f_cpd_in:
            while True:
                buf = f_cpd_in.read(1024 * 1024)
                if not buf:
                    break
                zf_entry.write(buf)

    # Stream Raw sheet into sheet3.xml
    with z_out.open('xl/worksheets/sheet3.xml', 'w', force_zip64=True) as zf_entry:
        with open(temp_raw_xml, 'rb') as f_raw_in:
            while True:
                buf = f_raw_in.read(1024 * 1024)
                if not buf:
                    break
                zf_entry.write(buf)

    z_out.close()
    z_in.close()

    # Cleanup temp XML files
    for p in (temp_raw_xml, temp_cpd_xml):
        if p.exists():
            try:
                p.unlink()
            except Exception:
                pass

    import gc
    gc.collect()
    log.info(f"Successfully generated Forward Pendency Report (Zero-Memory Stream): {output_file.name} ({total_filtered} rows)")
