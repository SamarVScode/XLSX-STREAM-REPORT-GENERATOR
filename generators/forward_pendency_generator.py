#!/usr/bin/env python3
"""
Forward Pendency Report Generator Module for ei_stream_server
=============================================================
Reads 'raw_data_North' from input Excel file, filters rows where Source_DC is in allowed list,
computes the 'Aging Category' column right beside 'Aging', and generates output workbook:
  1. Summary Sheet (3 Sidewise Tables with Red/Green highlights)
  2. CPD-DID pendency Sheet (P0, P1, P2 & P3 actual row details)
  3. RAW Sheet (Full filtered rows dataset with Aging Category)

Uses Single-Pass Zero-Memory Streaming Engine (core.stream_engine):
- O(1) Memory Footprint (< 35MB RAM)
- Direct XML disk streaming for massive datasets
"""

import sys
import re
import logging
from pathlib import Path
from datetime import datetime, date, timedelta
from collections import defaultdict
import openpyxl
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# Ensure server root is in sys.path
SERVER_ROOT = Path(__file__).resolve().parent.parent
if str(SERVER_ROOT) not in sys.path:
    sys.path.insert(0, str(SERVER_ROOT))

try:
    from config.dc_config import ALLOWED_SOURCE_DCS, ALLOWED_DCS_SET, ALLOWED_DCS_SET_LOWER, normalize_dc_code, is_allowed_dc
except ImportError:
    from dc_config import ALLOWED_SOURCE_DCS, ALLOWED_DCS_SET, ALLOWED_DCS_SET_LOWER, normalize_dc_code, is_allowed_dc

from core.stream_engine import (
    XmlSheetWriter,
    assemble_stream_workbook,
    open_stream_reader,
    get_sheet_names,
    ColumnFinder
)

PRIMARY_NORTH_DCS = ['ALG', 'AYP', 'DEO', 'JHS', 'JNP', 'MAU', 'MRZ', 'MTH', 'MZN', 'RBR', 'SPR']
AGING_CATEGORIES = ['0-2 days', '3-5 days', '5-10 days', '>10 days']
BASE_EXCEL_DATE = datetime(1899, 12, 30)

log = logging.getLogger("ei_stream_server.forward_pendency")


def extract_reference_date(input_path: Path) -> date:
    """Extracts report date from filename or defaults to current date."""
    fname = input_path.name
    # Try YYYY-Mon-DD, e.g., 2026-Oct-07
    m = re.search(r'(\d{4})-([A-Za-z]{3})-(\d{2})', fname)
    if m:
        try:
            return datetime.strptime(f"{m.group(1)}-{m.group(2)}-{m.group(3)}", "%Y-%b-%d").date()
        except ValueError:
            pass
    # Try YYYY-MM-DD
    m2 = re.search(r'(\d{4})-(\d{2})-(\d{2})', fname)
    if m2:
        try:
            return date(int(m2.group(1)), int(m2.group(2)), int(m2.group(3)))
        except ValueError:
            pass
    # Try DD-Mon-YYYY
    m3 = re.search(r'(\d{2})-([A-Za-z]{3})-(\d{4})', fname)
    if m3:
        try:
            return datetime.strptime(f"{m3.group(1)}-{m3.group(2)}-{m3.group(3)}", "%d-%b-%Y").date()
        except ValueError:
            pass
    return datetime.now().date()


def extract_cpd_date(val):
    """Parses date from Excel serial float, datetime, or text string."""
    if val is None:
        return None
    if isinstance(val, (datetime, date)):
        return val.date() if isinstance(val, datetime) else val
    if isinstance(val, (int, float)):
        try:
            return (BASE_EXCEL_DATE + timedelta(days=float(val))).date()
        except Exception:
            return None
    s = str(val).strip()
    if not s:
        return None
    for fmt in ('%Y-%m-%d %H:%M:%S', '%Y-%m-%d', '%d-%b-%Y', '%d-%b-%Y %H:%M:%S', '%d/%m/%Y', '%m/%d/%Y'):
        try:
            return datetime.strptime(s, fmt).date()
        except ValueError:
            pass
    m = re.search(r'(\d{4})-(\d{2})-(\d{2})', s)
    if m:
        try:
            return date(int(m.group(1)), int(m.group(2)), int(m.group(3)))
        except ValueError:
            pass
    return None


def compute_shipment_priority(cpd_val, target_date: date) -> str:
    """Classifies shipment based on CPD date: CPD, DID, or Non-CPD."""
    d = extract_cpd_date(cpd_val)
    if d is None:
        return "Non-CPD"
    if d == target_date:
        return "CPD"
    elif d < target_date:
        return "DID"
    else:
        return "Non-CPD"


def format_excel_datetime(val):
    """Converts Excel serial floats or datetimes into readable YYYY-MM-DD HH:MM:SS strings."""
    if val is None or val == "":
        return ""
    if isinstance(val, (datetime, date)):
        return val.strftime('%Y-%m-%d %H:%M:%S' if isinstance(val, datetime) and (val.hour or val.minute or val.second) else '%Y-%m-%d')
    if isinstance(val, (int, float)):
        if 30000 <= val <= 65000:
            try:
                dt = BASE_EXCEL_DATE + timedelta(days=float(val))
                if dt.microsecond >= 500000:
                    dt = dt + timedelta(seconds=1)
                dt = dt.replace(microsecond=0)
                if dt.hour or dt.minute or dt.second:
                    return dt.strftime('%Y-%m-%d %H:%M:%S')
                else:
                    return dt.strftime('%Y-%m-%d')
            except Exception:
                return str(val)
    return val


def compute_aging_category(val) -> str:
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


def normalize_priority(val) -> str:
    if val is None:
        return "Unknown"
    s = str(val).strip().upper()
    if not s:
        return "Unknown"
    if s in ("P0", "0", "0.0", "P-0", "P 0", "PRIORITY 0", "PRIORITY-0", "PRIORITY_0"):
        return "P0"
    if s in ("P1", "1", "1.0", "P-1", "P 1", "PRIORITY 1", "PRIORITY-1", "PRIORITY_1"):
        return "P1"
    if s in ("P2", "2", "2.0", "P-2", "P 2", "PRIORITY 2", "PRIORITY-2", "PRIORITY_2"):
        return "P2"
    if s in ("P3", "3", "3.0", "P-3", "P 3", "PRIORITY 3", "PRIORITY-3", "PRIORITY_3"):
        return "P3"
    if s in ("P4", "4", "4.0", "P-4", "P 4", "PRIORITY 4", "PRIORITY-4", "PRIORITY_4"):
        return "P4"
    if "P0" in s:
        return "P0"
    if "P1" in s:
        return "P1"
    if "P2" in s or "DID" in s:
        return "P2"
    if "P3" in s or "CPD" in s:
        return "P3"
    if "P4" in s:
        return "P4"
    return s


def write_side_table(ws, start_col: int, start_row: int, title: str, headers: list, data_matrix: list):
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
                if title == "CPD Pendency" and h_name in ['CPD', 'DID']:
                    cell.fill = red_fill
                    cell.font = red_font
                    highlighted = True
                elif title == "Aging wise report" and h_name in ['3-5 days', '5-10 days', '>10 days']:
                    cell.fill = red_fill
                    cell.font = red_font
                    highlighted = True
                elif title == "Priority Table":
                    if h_name in ['P0', 'P1', 'P2', 'P3']:
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


def build_summary_sheet_from_pivots(out_wb, t_cpd_pivot, t_aging_pivot, t_prio_pivot):
    """Generates the Summary tab with Table 1 (CPD Pendency), Table 2 (Aging wise), and Table 3 (Priority)."""
    ws = out_wb.active
    ws.title = "Summary"
    ws.sheet_view.showGridLines = False

    dc_list = list(PRIMARY_NORTH_DCS)
    for dc in ALLOWED_SOURCE_DCS:
        if dc.upper() not in dc_list:
            if any(t_cpd_pivot[dc.upper()].values()) or any(t_aging_pivot[dc.upper()].values()) or any(t_prio_pivot[dc.upper()].values()):
                dc_list.append(dc.upper())

    # 1. CPD Pendency Table (FIRST)
    t1_headers = ["Source DC", "CPD", "DID", "Total Pendency"]
    t1_data = []
    tot_cpd = 0
    tot_did = 0
    for dc in dc_list:
        c_cpd = t_cpd_pivot[dc]["CPD"]
        c_did = t_cpd_pivot[dc]["DID"]
        t1_data.append([dc, c_cpd, c_did, c_cpd + c_did])
        tot_cpd += c_cpd
        tot_did += c_did
    t1_data.append(["Total", tot_cpd, tot_did, tot_cpd + tot_did])

    # 2. Aging wise report (SECOND)
    t2_headers = ["Source DC"] + AGING_CATEGORIES + ["Total Pendency"]
    t2_data = []
    tot_cats = defaultdict(int)
    for dc in dc_list:
        row_vals = [dc]
        row_tot = 0
        for cat in AGING_CATEGORIES:
            cnt = t_aging_pivot[dc][cat]
            row_vals.append(cnt)
            row_tot += cnt
            tot_cats[cat] += cnt
        row_vals.append(row_tot)
        t2_data.append(row_vals)
    t2_data.append(["Total"] + [tot_cats[c] for c in AGING_CATEGORIES] + [sum(tot_cats.values())])

    # 3. Priority Table (THIRD)
    prio_keys = ["P0", "P1", "P2", "P3", "P4"]
    t3_headers = ["Source DC"] + prio_keys + ["Total Pendency"]
    t3_data = []
    tot_prios = defaultdict(int)
    for dc in dc_list:
        row_vals = [dc]
        row_tot = 0
        for p in prio_keys:
            cnt = t_prio_pivot[dc][p]
            row_vals.append(cnt)
            row_tot += cnt
            tot_prios[p] += cnt
        row_vals.append(row_tot)
        t3_data.append(row_vals)
    t3_data.append(["Total"] + [tot_prios[p] for p in prio_keys] + [sum(tot_prios.values())])

    start_row = 2
    # Table 1: CPD Pendency (cols 2 to 5 -> B to E)
    write_side_table(ws, start_col=2,  start_row=start_row, title="CPD Pendency",     headers=t1_headers, data_matrix=t1_data)
    # Gap col F (col 6)
    # Table 2: Aging wise report (cols 7 to 12 -> G to L)
    write_side_table(ws, start_col=7,  start_row=start_row, title="Aging wise report", headers=t2_headers, data_matrix=t2_data)
    # Gap col M (col 13)
    # Table 3: Priority Table (cols 14 to 20 -> N to T)
    write_side_table(ws, start_col=14, start_row=start_row, title="Priority Table",    headers=t3_headers, data_matrix=t3_data)

    ws.column_dimensions['A'].width = 3
    ws.column_dimensions['F'].width = 4
    ws.column_dimensions['M'].width = 4

    for col in ws.columns:
        col_letter = get_column_letter(col[0].column)
        if col_letter in ['A', 'F', 'M']:
            continue
        max_len = max(len(str(cell.value or '')) for cell in col)
        ws.column_dimensions[col_letter].width = max(max_len + 3, 13)


def generate_forward_pendency_report(input_file: Path, output_file: Path):
    input_path = Path(input_file)
    output_path = Path(output_file)
    if not input_path.exists():
        raise FileNotFoundError(f"Input file not found: {input_path}")

    log.info(f"Loading input workbook for Forward Pendency Report (Single-Pass Stream): {input_path}")

    target_date = extract_reference_date(input_path)
    log.info(f"Target reference date for CPD/DID classification: {target_date}")

    all_sheets = get_sheet_names(input_path)
    sheet_map = {name.lower(): name for name in all_sheets}
    target_sheet = None
    for candidate in ['raw_data_north', 'raw_data', 'raw', 'praw data', 'data', 'north']:
        if candidate in sheet_map:
            target_sheet = sheet_map[candidate]
            break
    if not target_sheet and all_sheets:
        target_sheet = all_sheets[0]

    log.info(f"Using sheet '{target_sheet}' for Forward Pendency Report.")

    t_cpd_pivot = defaultdict(lambda: defaultdict(int))
    t_aging_pivot = defaultdict(lambda: defaultdict(int))
    t_prio_pivot = defaultdict(lambda: defaultdict(int))

    total_filtered = 0
    cpd_count = 0

    with open_stream_reader(input_path, sheet_name=target_sheet) as (headers, row_iter):
        if not headers:
            raise ValueError(f"Sheet '{target_sheet}' is empty or not found.")

        cf = ColumnFinder(headers, {
            'aging': ['aging', 'agingbucket', 'agebucket', 'ageing', 'agingdays'],
            'sdc': ['sourcedc', 'sourcedccode', 'sourcedcname', 'sourcehub', 'origin', 'origindc', 'source_dc'],
            'prio': ['customerpriorityv2', 'customerpriority', 'custpriorityv2', 'priority', 'prio', 'customer_priority_v2'],
            'shipment': ['pendingshipments', 'trackingno', 'waybill', 'trackingid', 'shipment', 'awb', 'tracking_number'],
            'attempt': ['attemptstatus', 'attempt', 'lateststatus', 'laststatus', 'deliveryattempt', 'attempt_status'],
            'cpd': ['cpdvalue', 'cpd_value', 'cpd']
        })

        aging_col_idx = cf.get('aging', 20)
        sdc_idx = cf.get('sdc', 15)
        dc_idx = cf.find(['dc', 'hubdc', 'facility', 'destinationdc'], default=-1)
        prio_idx = cf.get('prio', 13)
        shipment_idx = cf.get('shipment', 1)
        attempt_idx = cf.get('attempt', 23)
        cpd_val_idx = cf.get('cpd', 6)

        raw_header_out = list(headers)
        insert_cpd_pos = cpd_val_idx + 1
        raw_header_out.insert(insert_cpd_pos, "Shipment Priority")

        adjusted_aging_pos = (aging_col_idx + 1) if aging_col_idx < cpd_val_idx else (aging_col_idx + 2)
        raw_header_out.insert(adjusted_aging_pos, "Aging Category")

        cpd_headers = [
            "PendingShipments",
            "Source_DC",
            "DC",
            "Aging Category",
            "Attempt_Status",
            "CustomerPriorityV2",
            "Shipment Priority"
        ]

        cpd_writer = XmlSheetWriter("CPD-DID pendency", cpd_headers)
        raw_writer = XmlSheetWriter("RAW", raw_header_out)

        with cpd_writer, raw_writer:
            for row in row_iter:
                if not row or len(row) <= sdc_idx:
                    continue
                raw_sdc = row[sdc_idx]
                if raw_sdc is None:
                    continue
                sdc_upper = normalize_dc_code(raw_sdc)

                if is_allowed_dc(sdc_upper):
                    total_filtered += 1
                    aging_val = row[aging_col_idx] if len(row) > aging_col_idx else None
                    aging_cat = compute_aging_category(aging_val)

                    raw_prio = row[prio_idx] if len(row) > prio_idx else None
                    prio = normalize_priority(raw_prio)

                    cpd_val = row[cpd_val_idx] if len(row) > cpd_val_idx else None
                    ship_prio = compute_shipment_priority(cpd_val, target_date)

                    # Aggregate pivots
                    t_aging_pivot[sdc_upper][aging_cat] += 1
                    t_prio_pivot[sdc_upper][prio] += 1
                    if ship_prio in ("CPD", "DID"):
                        t_cpd_pivot[sdc_upper][ship_prio] += 1

                    # Write RAW row (format dates and insert computed columns)
                    r_out = list(row)
                    r_out[sdc_idx] = sdc_upper
                    r_out[cpd_val_idx] = format_excel_datetime(cpd_val)
                    if len(r_out) > 9 and r_out[9]:
                        r_out[9] = format_excel_datetime(r_out[9])
                    if len(r_out) > 11 and r_out[11]:
                        r_out[11] = format_excel_datetime(r_out[11])

                    r_out.insert(insert_cpd_pos, ship_prio)
                    r_out.insert(adjusted_aging_pos, aging_cat)
                    raw_writer.write_row(r_out)

                    # Write CPD-DID row if CPD, DID, or P0/P1
                    is_cpd_did = ship_prio in ("CPD", "DID")
                    is_p0_p1 = prio in ("P0", "P1")
                    if is_cpd_did or is_p0_p1:
                        cpd_count += 1
                        shipment = row[shipment_idx] if len(row) > shipment_idx and row[shipment_idx] is not None else ""
                        dc_val = str(row[dc_idx]).strip() if (0 <= dc_idx < len(row) and row[dc_idx] is not None) else ""
                        attempt_stat = row[attempt_idx] if len(row) > attempt_idx and row[attempt_idx] is not None else ""
                        cpd_writer.write_row([shipment, sdc_upper, dc_val, aging_cat, attempt_stat, prio, ship_prio])

    log.info(f"Filtered {total_filtered} matching rows ({cpd_count} CPD-DID rows).")

    # Build Summary sheet with pre-registered placeholder tabs for clean OpenXML assembly
    out_wb = Workbook()
    build_summary_sheet_from_pivots(out_wb, t_cpd_pivot, t_aging_pivot, t_prio_pivot)
    out_wb.create_sheet("CPD-DID pendency")
    out_wb.create_sheet("RAW")

    # Assemble final .xlsx
    assemble_stream_workbook(out_wb, [cpd_writer, raw_writer], output_path)
    log.info(f"Successfully generated Forward Pendency Report: {output_file.name} ({total_filtered} rows)")
