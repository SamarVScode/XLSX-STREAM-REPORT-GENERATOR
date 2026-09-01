#!/usr/bin/env python3
"""
Untraceable Report Generator Module for ei_report_server
========================================================
Reads 'Raw' sheet from Untraceable Report file, filters rows where Source DC is in allowed list,
computes shipment count pivot by Source DC and Age Bucket, and generates output workbook:
  1. Summary Sheet (Merged 'Age Bucket' super-header, 0-value DCs & 0-cells omitted, soft red highlight for >= 6-10 days)
  2. Raw Sheet (Filtered raw records for allowed DC Config hubs)
"""

import sys
import logging
from pathlib import Path
import pandas as pd
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# Ensure current directory is in sys.path for dc_config import
CURRENT_DIR = Path(__file__).resolve().parent
if str(CURRENT_DIR) not in sys.path:
    sys.path.insert(0, str(CURRENT_DIR))

# Import server ALLOWED_DCS_SET_LOWER from config.dc_config
try:
    from config.dc_config import ALLOWED_DCS_SET_LOWER
except ImportError:
    try:
        from dc_config import ALLOWED_DCS_SET_LOWER
    except ImportError:
        ALLOWED_DCS_SET_LOWER = {'alg', 'ayp', 'deo', 'jhs', 'jnp', 'knp', 'mau', 'mrz', 'mth', 'mzn', 'rbr', 'spr', 'vns', 'all'}

log = logging.getLogger("ei_stream_server.untraceable_report")

STANDARD_AGE_BUCKETS = [
    '0-2 Days', '3-5 Days', '6-10 Days', '11-20 Days', '21-30 Days', '>30 Days'
]

HIGHLIGHT_AGING_BUCKETS = {'6-10 Days', '11-20 Days', '21-30 Days', '>30 Days', '7-15 Days', '15-30 Days', '30-45 Days', '45+ Days'}

def generate_untraceable_report(input_file: Path, output_file: Path):
    log.info(f"Loading input workbook for Untraceable Report: {input_file}")
    
    input_path = str(input_file)
    output_path = str(output_file)

    # Read ONLY 'Raw' sheet (try calamine / openpyxl / pyxlsb / default)
    df_raw = None
    engines = ['calamine', 'openpyxl', 'pyxlsb', None] if not str(input_path).endswith('.xlsb') else ['pyxlsb', 'calamine', None]
    
    for eng in engines:
        try:
            df_raw = pd.read_excel(input_path, sheet_name='Raw', engine=eng) if eng else pd.read_excel(input_path, sheet_name='Raw')
            break
        except Exception:
            continue

    if df_raw is None:
        try:
            df_raw = pd.read_excel(input_path, sheet_name=0)
        except Exception as e:
            raise ValueError(f"Could not read spreadsheet '{input_path}': {e}")

    # Convert numeric fields
    if 'Amount' in df_raw.columns:
        df_raw['Amount'] = pd.to_numeric(df_raw['Amount'], errors='coerce').fillna(0)
    if 'ShipmentAmount' in df_raw.columns:
        df_raw['ShipmentAmount'] = pd.to_numeric(df_raw['ShipmentAmount'], errors='coerce').fillna(0)

    # Identify Source DC column
    source_dc_col = None
    for col in df_raw.columns:
        if str(col).strip().lower() in ('source dc', 'sourcedc', 'source_dc', 'dc'):
            source_dc_col = col
            break
    if not source_dc_col:
        source_dc_col = 'Source DC'

    # Filter raw data for DC Config hubs only
    df_filtered = df_raw[df_raw[source_dc_col].astype(str).str.strip().str.lower().isin(ALLOWED_DCS_SET_LOWER)].copy()
    log.info(f"Filtered {len(df_filtered)} records out of {len(df_raw)} total records based on dc_config.")

    # Identify Age Bucket column
    age_bucket_col = None
    for col in df_filtered.columns:
        if str(col).strip().lower() in ('age bucket', 'age_bucket', 'aging bucket', 'ageing bucket', 'age'):
            age_bucket_col = col
            break
    if not age_bucket_col:
        age_bucket_col = 'Age Bucket'
        if age_bucket_col not in df_filtered.columns:
            df_filtered[age_bucket_col] = '0-2 Days'

    # Determine age bucket order
    raw_buckets = df_filtered[age_bucket_col].dropna().unique().tolist()
    present_buckets = [b for b in STANDARD_AGE_BUCKETS if b in raw_buckets]
    other_buckets = [b for b in raw_buckets if b not in STANDARD_AGE_BUCKETS]
    all_buckets = present_buckets + sorted(other_buckets)
    if not all_buckets:
        all_buckets = STANDARD_AGE_BUCKETS

    # Pivot: Source DC x Age Buckets (Count using groupby.size().unstack())
    p_cnt = df_filtered.groupby([source_dc_col, age_bucket_col]).size().unstack(fill_value=0)
    p_cnt = p_cnt.reindex(columns=all_buckets, fill_value=0)

    # Calculate Sum of Amount for each Source DC
    amt_col_name = 'Amount' if 'Amount' in df_filtered.columns else ('ShipmentAmount' if 'ShipmentAmount' in df_filtered.columns else None)
    if amt_col_name:
        p_amt = df_filtered.groupby(source_dc_col)[amt_col_name].sum()
        p_cnt['Amount'] = p_amt.reindex(p_cnt.index, fill_value=0)
    else:
        p_cnt['Amount'] = 0

    # Filter out DCs with 0 total shipments
    p_cnt['Total_Shipments'] = p_cnt[all_buckets].sum(axis=1)
    p_cnt = p_cnt[p_cnt['Total_Shipments'] > 0].drop(columns=['Total_Shipments'])

    wb = openpyxl.Workbook()

    # Styling themes
    header_fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid") # Dark Navy
    alt_row_fill = PatternFill(start_color="F9FAFB", end_color="F9FAFB", fill_type="solid")
    white_fill = PatternFill(start_color="FFFFFF", end_color="FFFFFF", fill_type="solid")
    
    # Highlight fill for aging bucket 6-10 days or more with count > 0
    aging_highlight_fill = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid") # Soft Red/Pink
    font_aging_highlight = Font(name="Calibri", size=11, bold=True, color="9C0006") # Dark Red text

    font_header = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    font_regular = Font(name="Calibri", size=11)

    thin_border_side = Side(border_style="thin", color="D9D9D9")
    border_cell = Border(left=thin_border_side, right=thin_border_side, top=thin_border_side, bottom=thin_border_side)
    no_border = Border()

    # ----------------------------------------------------
    # Sheet 1: Summary (Outer Gridlines Removed, All Table Cells Have Borders)
    # ----------------------------------------------------
    ws_sum = wb.active
    ws_sum.title = "Summary"
    ws_sum.sheet_view.showGridLines = False

    num_buckets = len(all_buckets)
    age_start_col = 2
    age_end_col = 1 + num_buckets
    amt_col_idx = 2 + num_buckets

    ws_sum.row_dimensions[1].height = 22
    ws_sum.row_dimensions[2].height = 22

    # --- Header 1: Source DC (Merged A1:A2) ---
    ws_sum.merge_cells(start_row=1, start_column=1, end_row=2, end_column=1)
    ws_sum.cell(row=1, column=1, value="Source DC")
    for r in range(1, 3):
        c = ws_sum.cell(row=r, column=1)
        c.font = font_header
        c.fill = header_fill
        c.alignment = Alignment(horizontal="center", vertical="center")
        c.border = border_cell

    # --- Header 2: Age Bucket Super Header (Merged B1 to G1) ---
    ws_sum.merge_cells(start_row=1, start_column=age_start_col, end_row=1, end_column=age_end_col)
    ws_sum.cell(row=1, column=age_start_col, value="Age Bucket")
    for col_idx in range(age_start_col, age_end_col + 1):
        c = ws_sum.cell(row=1, column=col_idx)
        c.font = font_header
        c.fill = header_fill
        c.alignment = Alignment(horizontal="center", vertical="center")
        c.border = border_cell

    # Row 2 Sub-Headers for Age Buckets
    for idx, b_text in enumerate(all_buckets):
        col_idx = age_start_col + idx
        c = ws_sum.cell(row=2, column=col_idx, value=b_text)
        c.font = font_header
        c.fill = header_fill
        c.alignment = Alignment(horizontal="center", vertical="center")
        c.border = border_cell

    # --- Header 3: Amount (Merged H1:H2) ---
    ws_sum.merge_cells(start_row=1, start_column=amt_col_idx, end_row=2, end_column=amt_col_idx)
    ws_sum.cell(row=1, column=amt_col_idx, value="Amount")
    for r in range(1, 3):
        c = ws_sum.cell(row=r, column=amt_col_idx)
        c.font = font_header
        c.fill = header_fill
        c.alignment = Alignment(horizontal="center", vertical="center")
        c.border = border_cell

    # Write Table Data (Row 3 onwards, excluding Grand Total)
    row_idx = 3
    for dc_name, row_data in p_cnt.iterrows():
        c1 = ws_sum.cell(row=row_idx, column=1, value=str(dc_name))
        c1.font = font_regular
        c1.border = border_cell

        for col_idx, b_name in enumerate(list(all_buckets) + ["Amount"], start=2):
            val = row_data[b_name]
            cell_val = val if val > 0 else "" # Hide 0 values
            c = ws_sum.cell(row=row_idx, column=col_idx, value=cell_val)
            c.alignment = Alignment(horizontal="right", vertical="center")
            c.border = border_cell # Uniform border for all table cells
            
            if cell_val != "":
                c.number_format = "#,##0"

            # Highlight cells in aging bucket columns >= 6-10 days with count > 0
            if b_name in HIGHLIGHT_AGING_BUCKETS and val > 0:
                c.fill = aging_highlight_fill
                c.font = font_aging_highlight
            else:
                c.font = font_regular

        row_idx += 1

    # Auto-fit column widths for Summary
    for col in ws_sum.columns:
        max_len = 0
        col_letter = get_column_letter(col[0].column)
        for cell in col:
            val_str = str(cell.value or '')
            if cell.number_format and '#' in cell.number_format and isinstance(cell.value, (int, float)):
                val_str = f"{cell.value:,.0f}"
            max_len = max(max_len, len(val_str))
        ws_sum.column_dimensions[col_letter].width = max(max_len + 4, 12)

    # ----------------------------------------------------
    # Sheet 2: Raw (Filtered Data for DC Config Hubs)
    # ----------------------------------------------------
    ws_raw = wb.create_sheet(title="Raw")
    ws_raw.sheet_view.showGridLines = True

    raw_headers = list(df_filtered.columns)
    for col_idx, h_text in enumerate(raw_headers, start=1):
        cell = ws_raw.cell(row=1, column=col_idx, value=h_text)
        cell.font = font_header
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = border_cell
    ws_raw.row_dimensions[1].height = 25

    for row_idx, row_values in enumerate(df_filtered.values, start=2):
        use_alt = (row_idx % 2 == 0)
        for col_idx, val in enumerate(row_values, start=1):
            col_name = raw_headers[col_idx - 1]
            cell = ws_raw.cell(row=row_idx, column=col_idx)
            
            if pd.isna(val):
                cell.value = ""
            elif isinstance(val, (int, float)):
                cell.value = val
                if col_name in ['Amount', 'ShipmentAmount']:
                    cell.number_format = "#,##0"
                    cell.alignment = Alignment(horizontal="right")
                elif col_name == 'CustomerPinCode':
                    cell.number_format = "0"
                    cell.alignment = Alignment(horizontal="center")
                else:
                    cell.alignment = Alignment(horizontal="right")
            else:
                cell.value = str(val)
                cell.alignment = Alignment(horizontal="left")

            cell.font = font_regular
            cell.border = border_cell
            if use_alt:
                cell.fill = alt_row_fill

    for col in ws_raw.columns:
        max_len = 0
        col_letter = get_column_letter(col[0].column)
        for cell in col:
            val_str = str(cell.value or '')
            if cell.number_format and '#' in cell.number_format and isinstance(cell.value, (int, float)):
                val_str = f"{cell.value:,.0f}"
            max_len = max(max_len, len(val_str))
        ws_raw.column_dimensions[col_letter].width = min(max(max_len + 3, 12), 40)

    # Save single output file
    wb.save(output_path)
    try:
        wb.close()
    except Exception:
        pass

    del wb
    import gc
    gc.collect()
    log.info(f"Successfully generated Untraceable Report: {output_path}")
    return output_path
