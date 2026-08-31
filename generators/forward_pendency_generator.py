#!/usr/bin/env python3
"""
Forward Pendency Report Generator Module for ei_report_server
=============================================================
Reads 'raw_data_North' from input Excel file, filters rows where Source_DC is in
{'alg', 'ayp', 'deo', 'jhs', 'jnp', 'mau', 'mrz', 'mth', 'mzn', 'rbr', 'spr'},
computes the 'Aging Category' column right beside 'Aging', and generates output workbook with:
  1. Summary Sheet (3 Sidewise Tables with Red/Green highlights)
  2. CPD-DID pendency Sheet (P2 & P3 actual row details)
  3. RAW Sheet (Full 6,775 filtered rows dataset)
"""

import logging
from pathlib import Path
from collections import defaultdict
import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter
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

def build_summary_sheet(out_wb, filtered_data_rows, aging_cat_idx, sdc_idx, prio_idx):
    """Generates the Summary tab with Table 1, Table 2, and Table 3 placed SIDEWISE."""
    ws = out_wb.create_sheet(title="Summary", index=0)
    ws.sheet_view.showGridLines = False

    t1_pivot = defaultdict(lambda: defaultdict(int))
    t2_pivot = defaultdict(lambda: defaultdict(int))
    t3_pivot = defaultdict(lambda: defaultdict(int))

    for r in filtered_data_rows[1:]:
        sdc = str(r[sdc_idx]).upper()
        cat = str(r[aging_cat_idx])
        prio = str(r[prio_idx]) if r[prio_idx] is not None else "Unknown"

        t1_pivot[sdc][cat] += 1
        t2_pivot[sdc][prio] += 1

        if prio == "P3":
            t3_pivot[sdc]["CPD (P3)"] += 1
        elif prio == "P2":
            t3_pivot[sdc]["DID (P2)"] += 1

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

def build_cpd_did_sheet(out_wb, filtered_data_rows, aging_cat_idx, sdc_idx, prio_idx, shipment_idx, attempt_idx):
    """Generates the 'CPD-DID' tab containing all actual P2 and P3 rows."""
    ws = out_wb.create_sheet(title="CPD-DID", index=1)

    headers = ["PendingShipments", "Source_DC", "Aging Category", "Attempt_Status", "CustomerPriorityV2"]
    header_fill = PatternFill("solid", fgColor="312E81")
    header_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    data_font = Font(name="Calibri", size=11, color="1F2937")
    thin_side = Side(style="thin", color="CBD5E1")
    border = Border(left=thin_side, right=thin_side, top=thin_side, bottom=thin_side)
    center_align = Alignment(horizontal="center", vertical="center")
    left_align = Alignment(horizontal="left", vertical="center")

    for c_idx, h_text in enumerate(headers, start=1):
        cell = ws.cell(row=1, column=c_idx, value=h_text)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = center_align
        cell.border = border

    current_row = 2
    for r in filtered_data_rows[1:]:
        prio = str(r[prio_idx]) if r[prio_idx] is not None else ""
        if prio in ("P2", "P3"):
            shipment = r[shipment_idx]
            sdc = str(r[sdc_idx]).upper() if r[sdc_idx] is not None else ""
            aging_cat = r[aging_cat_idx]
            attempt_stat = r[attempt_idx] if r[attempt_idx] is not None else ""

            row_vals = [shipment, sdc, aging_cat, attempt_stat, prio]
            for c_idx, val in enumerate(row_vals, start=1):
                cell = ws.cell(row=current_row, column=c_idx, value=val)
                cell.font = data_font
                cell.border = border
                cell.alignment = left_align if c_idx == 1 else center_align
            current_row += 1

    for col in ws.columns:
        col_letter = get_column_letter(col[0].column)
        max_len = max(len(str(cell.value or '')) for cell in col)
        ws.column_dimensions[col_letter].width = max(max_len + 4, 16)

def generate_forward_pendency_report(input_file: Path, output_file: Path):
    """
    Main generator pipeline for Forward Pendency Report.
    """
    if not Path(input_file).exists():
        raise FileNotFoundError(f"Input file not found: {input_file}")

    log.info(f"Loading input workbook for Forward Pendency Report: {input_file}")
    in_wb = openpyxl.load_workbook(input_file, data_only=True, read_only=True)

    # Sheet selection with fallback: 'raw_data_North' -> 'raw_data' (with case-insensitive fallback)
    target_sheet = None
    if 'raw_data_North' in in_wb.sheetnames:
        target_sheet = 'raw_data_North'
    elif 'raw_data' in in_wb.sheetnames:
        target_sheet = 'raw_data'
    else:
        sheet_map = {name.lower(): name for name in in_wb.sheetnames}
        if 'raw_data_north' in sheet_map:
            target_sheet = sheet_map['raw_data_north']
        elif 'raw_data' in sheet_map:
            target_sheet = sheet_map['raw_data']

    if not target_sheet:
        raise ValueError(f"Neither 'raw_data_North' nor 'raw_data' sheet found in input. Available: {in_wb.sheetnames}")

    in_ws = in_wb[target_sheet]
    log.info(f"Reading rows from '{target_sheet}' sheet...")

    filtered_rows = []
    header_list = None
    aging_col_idx = 20
    sdc_idx = 15
    prio_idx = 13
    shipment_idx = 1
    attempt_idx = 23
    total_processed = 0

    for i, row in enumerate(in_ws.iter_rows(values_only=True)):
        if i == 0:
            header_list = list(row)
            if 'Aging' in header_list:
                aging_col_idx = header_list.index('Aging')
            if 'Source_DC' in header_list:
                sdc_idx = header_list.index('Source_DC')
            if 'CustomerPriorityV2' in header_list:
                prio_idx = header_list.index('CustomerPriorityV2')
            if 'PendingShipments' in header_list:
                shipment_idx = header_list.index('PendingShipments')
            if 'Attempt_Status' in header_list:
                attempt_idx = header_list.index('Attempt_Status')

            header_list.insert(aging_col_idx + 1, 'Aging Category')
            filtered_rows.append(header_list)

            sdc_idx_new = sdc_idx + 1 if sdc_idx > aging_col_idx else sdc_idx
            prio_idx_new = prio_idx + 1 if prio_idx > aging_col_idx else prio_idx
            shipment_idx_new = shipment_idx + 1 if shipment_idx > aging_col_idx else shipment_idx
            attempt_idx_new = attempt_idx + 1 if attempt_idx > aging_col_idx else attempt_idx
            aging_cat_idx = aging_col_idx + 1
            continue

        total_processed += 1
        raw_sdc = row[sdc_idx] if len(row) > sdc_idx else None
        sdc = str(raw_sdc).strip().lower() if raw_sdc is not None else ''

        if sdc in ALLOWED_DCS_SET_LOWER:
            row_list = list(row)
            aging_val = row_list[aging_col_idx] if len(row_list) > aging_col_idx else None
            aging_cat = compute_aging_category(aging_val)
            row_list.insert(aging_col_idx + 1, aging_cat)
            filtered_rows.append(row_list)

    log.info(f"Processed {total_processed} source rows. Filtered {len(filtered_rows) - 1} matching rows.")
    in_wb.close()

    out_wb = openpyxl.Workbook()
    
    # 1. RAW Sheet (Main filtered data tab)
    data_ws = out_wb.active
    data_ws.title = "Raw"
    for r in filtered_rows:
        data_ws.append(r)

    # 2. Summary Sheet
    build_summary_sheet(out_wb, filtered_rows, aging_cat_idx, sdc_idx_new, prio_idx_new)

    # 3. CPD-DID Sheet
    build_cpd_did_sheet(out_wb, filtered_rows, aging_cat_idx, sdc_idx_new, prio_idx_new, shipment_idx_new, attempt_idx_new)

    out_wb.save(output_file)
    log.info(f"Successfully generated Forward Pendency Report: {output_file}")
