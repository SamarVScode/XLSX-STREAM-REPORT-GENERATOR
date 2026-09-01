#!/usr/bin/env python3
"""
EOB Report Generator Module for ei_report_server
=================================================
Reads 'Raw' (or 'raw_data') sheet from input Excel/XLSB file, filters rows where Source_DC is in
{'ALG', 'AYP', 'DEO', 'JNP', 'MAU', 'MRZ'}, and generates an output workbook with:
  1. Summary Sheet:
     - Table 1: Ageing Bucket Wise Priority Shipment Count per Source DC (Cells with count > 0 in Ageing Bucket columns highlighted in Red).
     - Table 2: Latest Status Count per Source DC (without 'UD' prefix) placed side-by-side.
     - Starts at Column A (left-most start) with exactly 1 empty column gap between Table 1 and Table 2.
     - Both tables start at Row 1 (no extra title banners or extra headers).
     - Worksheet gridlines disabled for empty background cells.
  2. raw_data Sheet:
     - Full filtered dataset containing all original columns for target Source DCs.
"""

import logging
from pathlib import Path
import pandas as pd
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# Try fast calamine engine first, fall back to pyxlsb / openpyxl
try:
    from python_calamine import CalamineWorkbook
    HAS_CALAMINE = True
except ImportError:
    HAS_CALAMINE = False

try:
    from pyxlsb import open_workbook as open_xlsb
    HAS_PYXLSB = True
except ImportError:
    HAS_PYXLSB = False

log = logging.getLogger("ei_stream_server.eob_generator")

try:
    from config.dc_config import ALLOWED_SOURCE_DCS
except ImportError:
    try:
        from dc_config import ALLOWED_SOURCE_DCS
    except ImportError:
        ALLOWED_SOURCE_DCS = ['ALG', 'AYP', 'DEO', 'JHS', 'JNP', 'KNP', 'MAU', 'MRZ', 'MTH', 'MZN', 'RBR', 'SPR', 'VNS']

# Dynamic list of target Source DCs from dc_config (excluding 'ALL')
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


def _resolve_raw_sheet_name(input_file_path: Path) -> str:
    """Find the raw data sheet name in the workbook ('Raw', 'raw_data', etc.)."""
    if input_file_path.suffix.lower() == '.xlsb':
        if HAS_CALAMINE:
            wb = CalamineWorkbook.from_path(str(input_file_path))
            sheet_names = wb.sheet_names
        elif HAS_PYXLSB:
            with open_xlsb(str(input_file_path)) as wb:
                sheet_names = wb.sheets
        else:
            sheet_names = ['Raw', 'raw_data']
    else:
        wb = openpyxl.load_workbook(input_file_path, read_only=True)
        sheet_names = wb.sheetnames
        wb.close()

    sheet_map = {name.lower(): name for name in sheet_names}
    for candidate in ['raw', 'raw_data', 'raw_data_north', 'praw data']:
        if candidate in sheet_map:
            return sheet_map[candidate]
            
    return sheet_names[0] if sheet_names else 'Raw'


def read_raw_data(input_file_path: Path) -> pd.DataFrame:
    """Reads the raw data sheet from input file."""
    sheet_name = _resolve_raw_sheet_name(input_file_path)
    ext = input_file_path.suffix.lower()
    
    if HAS_CALAMINE:
        try:
            wb = CalamineWorkbook.from_path(str(input_file_path))
            sheet = wb.get_sheet_by_name(sheet_name)
            data = sheet.to_python()
            if data:
                header = [str(c).strip() if c is not None else f"Unnamed_{i}" for i, c in enumerate(data[0])]
                return pd.DataFrame(data[1:], columns=header)
        except Exception:
            pass

    if ext == '.xlsb':
        if HAS_PYXLSB:
            data = []
            with open_xlsb(str(input_file_path)) as wb:
                with wb.get_sheet(sheet_name) as sheet:
                    for row in sheet.rows():
                        data.append([cell.v for cell in row])
            if data:
                header = [str(c).strip() if c is not None else f"Unnamed_{i}" for i, c in enumerate(data[0])]
                return pd.DataFrame(data[1:], columns=header)
        try:
            return pd.read_excel(input_file_path, sheet_name=sheet_name, engine='pyxlsb')
        except Exception:
            return pd.read_excel(input_file_path, sheet_name=sheet_name)
    else:
        try:
            return pd.read_excel(input_file_path, sheet_name=sheet_name)
        except Exception:
            return pd.read_excel(input_file_path, sheet_name=0)


def get_short_status(status_str):
    if not status_str or pd.isna(status_str):
        return 'Unknown'
    s = str(status_str).strip()
    return STATUS_SHORTFORMS.get(s, s)


def generate_eob_report(input_file: Path, output_file: Path):
    """Main generator function for EOB Priority & Status Report."""
    input_path = Path(input_file)
    output_path = Path(output_file)

    if not input_path.exists():
        raise FileNotFoundError(f"Input file not found: {input_file}")

    log.info(f"Generating EOB Priority & Status Report for: {input_path.name}")
    df_raw = read_raw_data(input_path)

    # Clean up empty trailing rows
    if 'Tracking No' in df_raw.columns:
        df_raw = df_raw[df_raw['Tracking No'].notnull() & (df_raw['Tracking No'] != '')]
    elif 'tracking_id' in df_raw.columns:
        df_raw = df_raw[df_raw['tracking_id'].notnull() & (df_raw['tracking_id'] != '')]
    else:
        df_raw = df_raw.dropna(how='all')

    # Detect Source DC column name
    sdc_col = None
    for candidate in ['Source DC', 'Source_DC', 'source_dc', 'dc']:
        if candidate in df_raw.columns:
            sdc_col = candidate
            break

    if not sdc_col:
        raise ValueError(f"Could not find 'Source DC' column in raw sheet. Available columns: {list(df_raw.columns)}")

    # Filter for target Source DCs (case-insensitive)
    df_raw['Source DC Clean'] = df_raw[sdc_col].astype(str).str.strip().str.upper()
    filtered_df = df_raw[df_raw['Source DC Clean'].isin(TARGET_SOURCE_DCS)].copy()

    log.info(f"Filtered {len(filtered_df)} total records for Source DCs: {TARGET_SOURCE_DCS}")

    # Detect Latest Status column
    status_col = None
    for candidate in ['Latest Status', 'Latest_Status', 'status_status', 'status']:
        if candidate in df_raw.columns:
            status_col = candidate
            break

    if not status_col:
        status_col = 'Latest Status'
        filtered_df['Latest Status'] = 'Unknown'

    # Detect Ageing Bucket column
    ageing_col = None
    for candidate in ['Ageing Bucket', 'Ageing_Bucket', 'Aging Bucket', 'aging_bucket']:
        if candidate in df_raw.columns:
            ageing_col = candidate
            break

    if not ageing_col:
        ageing_col = 'Ageing Bucket'
        filtered_df['Ageing Bucket'] = 'Unknown'

    # Map Latest Status to shortform
    filtered_df['Latest Status Short'] = filtered_df[status_col].map(get_short_status)

    # Prepare raw export dataframe
    raw_export_df = filtered_df.drop(columns=['Source DC Clean', 'Latest Status Short'])

    # --- PIVOT TABLE 1: Source DC (Rows) vs Ageing Bucket (Cols) ---
    ageing_order = ['1-2 days', '3-5 days', '6-10 days', '11-15 days', '16-20 days', '>20 days']
    existing_buckets = filtered_df[ageing_col].dropna().unique().tolist()
    sorted_buckets = [b for b in ageing_order if b in existing_buckets]
    for b in sorted(existing_buckets):
        if b not in sorted_buckets:
            sorted_buckets.append(b)

    t1_df = pd.crosstab(filtered_df['Source DC Clean'], filtered_df[ageing_col])
    t1_df = t1_df.reindex(index=TARGET_SOURCE_DCS, columns=sorted_buckets, fill_value=0)
    t1_df['Grand Total'] = t1_df.sum(axis=1)

    t1_total_row = t1_df.sum(axis=0)
    t1_total_row.name = 'Grand Total'
    t1_df = pd.concat([t1_df, pd.DataFrame([t1_total_row])])

    # --- PIVOT TABLE 2: Source DC (Rows) vs Latest Status Shortform (Cols) ---
    t2_df = pd.crosstab(filtered_df['Source DC Clean'], filtered_df['Latest Status Short'])
    t2_statuses = sorted(t2_df.columns.tolist())
    t2_df = t2_df.reindex(index=TARGET_SOURCE_DCS, columns=t2_statuses, fill_value=0)
    t2_df['Grand Total'] = t2_df.sum(axis=1)

    t2_total_row = t2_df.sum(axis=0)
    t2_total_row.name = 'Grand Total'
    t2_df = pd.concat([t2_df, pd.DataFrame([t2_total_row])])

    # --- BUILD OPENPYXL WORKBOOK ---
    wb = openpyxl.Workbook()

    # 1. Summary Sheet
    ws_summary = wb.active
    ws_summary.title = "Summary"
    ws_summary.views.sheetView[0].showGridLines = False

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
    start_c1 = 1  # Column A

    # WRITE TABLE 1
    headers_t1 = ['Source DC'] + list(t1_df.columns)
    for c_idx, h in enumerate(headers_t1, start=start_c1):
        cell = ws_summary.cell(row=start_r, column=c_idx, value=h)
        cell.font = font_header
        cell.fill = fill_header
        cell.alignment = align_center
        cell.border = border_header

    for r_offset, (row_name, row_series) in enumerate(t1_df.iterrows(), start=start_r + 1):
        is_total = (row_name == 'Grand Total')
        c_cell = ws_summary.cell(row=r_offset, column=start_c1, value=row_name)
        c_cell.font = font_total if is_total else font_data
        c_cell.alignment = align_center
        c_cell.border = border_total if is_total else border_cell
        if is_total:
            c_cell.fill = fill_total

        for c_idx, (col_name, val) in enumerate(row_series.items(), start=start_c1 + 1):
            cell = ws_summary.cell(row=r_offset, column=c_idx, value=int(val))
            cell.alignment = align_right

            if is_total:
                cell.font = font_total
                cell.fill = fill_total
                cell.border = border_total
            else:
                cell.border = border_cell
                if col_name != 'Grand Total' and val > 0:
                    cell.fill = fill_priority
                    cell.font = font_priority
                else:
                    cell.font = font_data
                    if r_offset % 2 == 0:
                        cell.fill = fill_zebra

    # WRITE TABLE 2 (Side-by-Side)
    start_c2 = start_c1 + len(headers_t1) + 1  # 1 empty column gap
    headers_t2 = ['Source DC'] + list(t2_df.columns)

    for c_idx, h in enumerate(headers_t2, start=start_c2):
        cell = ws_summary.cell(row=start_r, column=c_idx, value=h)
        cell.font = font_header
        cell.fill = fill_header
        cell.alignment = align_center
        cell.border = border_header

    for r_offset, (row_name, row_series) in enumerate(t2_df.iterrows(), start=start_r + 1):
        is_total = (row_name == 'Grand Total')
        c_cell = ws_summary.cell(row=r_offset, column=start_c2, value=row_name)
        c_cell.font = font_total if is_total else font_data
        c_cell.alignment = align_center
        c_cell.border = border_total if is_total else border_cell
        if is_total:
            c_cell.fill = fill_total

        for c_idx, (col_name, val) in enumerate(row_series.items(), start=start_c2 + 1):
            cell = ws_summary.cell(row=r_offset, column=c_idx, value=int(val))
            cell.alignment = align_right

            if is_total:
                cell.font = font_total
                cell.fill = fill_total
                cell.border = border_total
            else:
                cell.font = font_data
                cell.border = border_cell
                if r_offset % 2 == 0:
                    cell.fill = fill_zebra

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
    ws_raw.views.sheetView[0].showGridLines = True

    ws_raw.append(list(raw_export_df.columns))
    for c_idx in range(1, len(raw_export_df.columns) + 1):
        cell = ws_raw.cell(row=1, column=c_idx)
        cell.font = font_header
        cell.fill = fill_header
        cell.alignment = align_center

    for row in raw_export_df.itertuples(index=False):
        ws_raw.append(list(row))

    # Save Output
    wb.save(output_path)
    try:
        wb.close()
    except Exception:
        pass

    del wb, raw_export_df
    import gc
    gc.collect()
    log.info(f"Successfully generated EOB Report: {output_path.name}")
