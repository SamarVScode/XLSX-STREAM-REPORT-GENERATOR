#!/usr/bin/env python3
"""
Conversion Report Generator Module for ei_report_server
========================================================
Reads 'E2E_DC' and 'Agent_view' tabs from input Excel file,
filters for allowed Source_DCs, computes percentages and color thresholds,
and writes formatted output workbook.
"""

import re
import logging
from datetime import datetime, date
from pathlib import Path
from typing import Optional
import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

try:
    from config.dc_config import ALLOWED_DCS_SET as ALLOWED_DCS
except ImportError:
    from dc_config import ALLOWED_DCS_SET as ALLOWED_DCS
PCT_COLS_DC = ['Succ_pickup%', 'Succ_del%', 'COD_Succ_del%', 'PP_Succ_del%']

HEADER_FILL = PatternFill(start_color='1F4E79', end_color='1F4E79', fill_type='solid')
HEADER_FONT = Font(bold=True, color='FFFFFF', size=11)
THIN_BORDER = Border(
    left=Side(style='thin'), right=Side(style='thin'),
    top=Side(style='thin'), bottom=Side(style='thin'),
)
CENTER = Alignment(horizontal='center', vertical='center')

FILL_RED = PatternFill(start_color='FFC7CE', end_color='FFC7CE', fill_type='solid')
FILL_YELLOW = PatternFill(start_color='FFEB9C', end_color='FFEB9C', fill_type='solid')
FILL_GREEN = PatternFill(start_color='C6EFCE', end_color='C6EFCE', fill_type='solid')
FONT_RED = Font(color='9C0006')
FONT_YELLOW = Font(color='9C6500')
FONT_GREEN = Font(color='006100')

log = logging.getLogger("ei_stream_server.conversion_report")

def _get_all_sheet_names(input_file: Path):
    """Return all sheet names from the workbook."""
    import openpyxl
    wb = openpyxl.load_workbook(input_file, read_only=True)
    names = wb.sheetnames
    wb.close()
    return names


def _resolve_sheet(input_file: Path, target: str) -> str:
    """Return the exact sheet name from the workbook that matches target case-insensitively."""
    names = _get_all_sheet_names(input_file)
    target_lower = target.lower()
    for name in names:
        if name.lower() == target_lower:
            return name
    raise ValueError(f"Sheet '{target}' not found in workbook. Available tabs: {names}")


def extract_report_date_from_agent_view(input_file: Path) -> Optional[str]:
    """
    Extracts the report date from the first column ('Date') of the 'Agent_view' tab.
    Returns normalized date string (e.g. '18-08-2026') or None if not found.
    """
    try:
        import openpyxl
        wb = openpyxl.load_workbook(str(input_file), read_only=True, data_only=True)
        target_name = None
        for name in wb.sheetnames:
            norm_name = name.strip().lower().replace(' ', '_').replace('-', '_')
            if 'agent_view' in norm_name or 'agentview' in norm_name:
                target_name = name
                break
        if not target_name:
            for name in wb.sheetnames:
                if 'agent' in name.strip().lower():
                    target_name = name
                    break
        if not target_name and len(wb.sheetnames) > 0:
            target_name = wb.sheetnames[0]

        ws = wb[target_name]
        date_col_idx = 0
        found_date_val = None

        for r_idx, row in enumerate(ws.iter_rows(values_only=True)):
            if r_idx == 0:
                for c_idx, h in enumerate(row):
                    if h is not None and 'date' in str(h).strip().lower():
                        date_col_idx = c_idx
                        break
            else:
                if len(row) > date_col_idx:
                    val = row[date_col_idx]
                    if val is not None and str(val).strip() not in ('', 'None', 'nan', 'NaT'):
                        found_date_val = val
                        break
                if found_date_val is None and len(row) > 0:
                    val = row[0]
                    if val is not None and str(val).strip() not in ('', 'None', 'nan', 'NaT'):
                        found_date_val = val
                        break
        wb.close()

        if found_date_val is not None:
            if isinstance(found_date_val, (datetime, date)):
                return found_date_val.strftime('%d-%m-%Y')
            val_str = str(found_date_val).strip()
            for fmt in ('%d-%m-%Y', '%Y-%m-%d', '%d/%m/%Y', '%Y/%m/%d', '%d-%b-%Y', '%d-%B-%Y', '%d-%m-%y', '%Y%m%d'):
                try:
                    parsed_dt = datetime.strptime(val_str[:10], fmt)
                    return parsed_dt.strftime('%d-%m-%Y')
                except Exception:
                    pass
            clean_str = re.sub(r'[^\w\-]', '-', val_str).strip('-')
            return clean_str if clean_str else None
    except Exception as ex:
        log.warning(f"Could not extract report date from Agent_view: {ex}")
    return None


def _clean_pct_val(val):
    """Convert string, float, or ratio percentage values to standard 0.0 - 1.0 float ratios."""
    if pd.isna(val):
        return 0.0
    had_pct = False
    if isinstance(val, str):
        val_str = val.strip()
        had_pct = '%' in val_str
        val_clean = val_str.rstrip('%').strip()
        try:
            val = float(val_clean)
        except ValueError:
            return 0.0
    else:
        try:
            val = float(val)
        except (ValueError, TypeError):
            return 0.0

    if had_pct or val > 1.0:
        return val / 100.0
    return val


def build_dc_view(input_file: Path) -> pd.DataFrame:
    sheet = _resolve_sheet(input_file, 'E2E_DC')
    df = pd.read_excel(input_file, sheet_name=sheet)
    dc_col = None
    for c in df.columns:
        if str(c).strip().lower() in ('source_dc', 'source dc', 'dc', 'sourcedc', 'hub'):
            dc_col = c
            break
    if dc_col and dc_col != 'Source_DC':
        df.rename(columns={dc_col: 'Source_DC'}, inplace=True)
    if 'Source_DC' in df.columns:
        df = df[df['Source_DC'].astype(str).str.strip().str.upper().isin(ALLOWED_DCS)].copy()

    if 'Picked-up' in df.columns:
        df['Picked-up'] = pd.to_numeric(df['Picked-up'], errors='coerce').fillna(0)
    if 'OFP' in df.columns:
        df['OFP'] = pd.to_numeric(df['OFP'], errors='coerce').fillna(0)

    for col in PCT_COLS_DC:
        if col in df.columns:
            df[col] = df[col].apply(_clean_pct_val)

    # Recalculate/ensure Succ_pickup% if missing or null
    if 'Picked-up' in df.columns and 'OFP' in df.columns:
        if 'Succ_pickup%' not in df.columns or df['Succ_pickup%'].sum() == 0:
            ofp_denom = df['OFP'].replace(0, float('nan'))
            df['Succ_pickup%'] = (df['Picked-up'] / ofp_denom).astype('float64').round(4).fillna(0.0)

    return df


def build_agent_view(input_file: Path) -> pd.DataFrame:
    sheet = _resolve_sheet(input_file, 'Agent_view')
    df = pd.read_excel(input_file, sheet_name=sheet)
    dc_col = None
    for c in df.columns:
        if str(c).strip().lower() in ('source_dc', 'source dc', 'dc', 'sourcedc', 'hub'):
            dc_col = c
            break
    if dc_col and dc_col != 'Source_DC':
        df.rename(columns={dc_col: 'Source_DC'}, inplace=True)
    if 'Source_DC' in df.columns:
        df = df[df['Source_DC'].astype(str).str.strip().str.upper().isin(ALLOWED_DCS)].copy()

    # Identify delivery update column name ('del_update' or 'del_ppdate')
    del_col = None
    if 'del_update' in df.columns:
        del_col = 'del_update'
    elif 'del_ppdate' in df.columns:
        del_col = 'del_ppdate'

    # Remove any pre-existing 'Total OFP' or 'Total OFD' columns
    for extra_col in ['Total OFP', 'Total OFD']:
        if extra_col in df.columns:
            df.drop(columns=[extra_col], inplace=True)

    # Coerce numeric columns
    cols_to_coerce = ['OFP', 'Picked-up', 'OFD']
    if del_col:
        cols_to_coerce.append(del_col)

    for col in cols_to_coerce:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)

    # REV % = Picked-up / OFP
    if 'OFP' in df.columns and 'Picked-up' in df.columns:
        ofp_denom = df['OFP'].replace(0, float('nan'))
        df['REV %'] = (df['Picked-up'] / ofp_denom).astype('float64').round(4).fillna(0.0)

    # FWD % = del_update / OFD (or del_ppdate / OFD)
    if 'OFD' in df.columns and del_col and del_col in df.columns:
        ofd_denom = df['OFD'].replace(0, float('nan'))
        df['FWD %'] = (df[del_col] / ofd_denom).astype('float64').round(4).fillna(0.0)

    return df


def _pct_fill_font(value, thresholds):
    lo, mid = thresholds
    if value < lo:
        return FILL_RED, FONT_RED
    elif value < mid:
        return FILL_YELLOW, FONT_YELLOW
    else:
        return FILL_GREEN, FONT_GREEN

def style_sheet(ws, pct_cols=None):
    for cell in ws[1]:
        cell.fill = HEADER_FILL
        cell.font = HEADER_FONT
        cell.alignment = CENTER
        cell.border = THIN_BORDER

    for row in ws.iter_rows(min_row=2, max_row=ws.max_row, max_col=ws.max_column):
        for cell in row:
            cell.border = THIN_BORDER
            if isinstance(cell.value, float):
                cell.number_format = '0.000'
            cell.alignment = Alignment(horizontal='center')

    if pct_cols:
        for row in ws.iter_rows(min_row=2, max_row=ws.max_row):
            for cell in row:
                if cell.column_letter in pct_cols and isinstance(cell.value, (int, float)):
                    thresholds = pct_cols[cell.column_letter]
                    fill, font = _pct_fill_font(cell.value, thresholds)
                    cell.fill = fill
                    cell.font = font
                    cell.number_format = '0.0%'

    ws.auto_filter.ref = ws.dimensions

def _col_letter_by_name(ws, col_name: str):
    """Return the column letter for a header cell matching col_name (case-insensitive)."""
    for cell in ws[1]:
        if cell.value and str(cell.value).strip().lower() == col_name.strip().lower():
            return cell.column_letter
    return None


def generate_conversion_report(input_file: Path, output_file: Path, sub_type: str = "Sameday"):
    """
    Main generator pipeline for Conversion Report.
    sub_type: 'Sameday' or 'D-1' — used to prefix sheet names and filename.
    """
    # Normalize label: 'sameday' -> 'Sameday', 'd1'/'d-1' -> 'D-1'
    label = sub_type.strip()
    if label.lower() in ("d1", "d-1"):
        label = "D-1"
    elif label.lower() == "sameday":
        label = "Sameday"

    log.info(f"Loading input workbook for Conversion Report [{label}]: {input_file}")
    dc_df    = build_dc_view(input_file)
    agent_df = build_agent_view(input_file)

    sheet_dc    = f"{label} DC_View"
    sheet_agent = f"{label} Agent_View"

    log.info(f"Writing outputs: {sheet_dc} ({len(dc_df)} rows), {sheet_agent} ({len(agent_df)} rows)")
    with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
        dc_df.to_excel(writer,    sheet_name=sheet_dc,    index=False)
        agent_df.to_excel(writer, sheet_name=sheet_agent, index=False)

    wb = load_workbook(output_file)
    for name in wb.sheetnames:
        ws = wb[name]
        if name == sheet_dc:
            dc_pct_cols = {}
            for col_name, thresholds in [
                ('Succ_pickup%', (0.75, 0.85)),
                ('Succ_del%', (0.80, 0.90)),
                ('COD_Succ_del%', (0.75, 0.85)),
                ('PP_Succ_del%', (0.87, 0.95))
            ]:
                col_let = _col_letter_by_name(ws, col_name)
                if col_let:
                    dc_pct_cols[col_let] = thresholds
                    log.info(f"{sheet_dc}: {col_name} mapped to column {col_let}")
            style_sheet(ws, pct_cols=dc_pct_cols if dc_pct_cols else None)
        elif name == sheet_agent:
            rev_col = _col_letter_by_name(ws, 'REV %')
            fwd_col = _col_letter_by_name(ws, 'FWD %')
            agent_pct_cols = {}
            if rev_col:
                agent_pct_cols[rev_col] = (0.75, 0.85)
                log.info(f"{sheet_agent}: REV % at column {rev_col}")
            if fwd_col:
                agent_pct_cols[fwd_col] = (0.80, 0.90)
                log.info(f"{sheet_agent}: FWD % at column {fwd_col}")
            style_sheet(ws, pct_cols=agent_pct_cols if agent_pct_cols else None)
        else:
            style_sheet(ws)
    wb.save(output_file)
    wb.close()
    log.info(f"Successfully generated Conversion Report [{label}]: {output_file}")
