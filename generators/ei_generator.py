#!/usr/bin/env python3
"""
EI Summary Report Generator Module for ei_stream_server
=======================================================
Processes 'Task_per_1k' and 'Raw' sheets to generate:
  1. SUMMARY sheet (Formatted Daily and WTD performance tables with color styling)
  2. Filtered sheet (Raw records filtered for target DCs)
  3. FWD EI sheet (Tracking starting with MYSC or MYSP)
  4. REVERSE EI sheet (Tracking starting with MYSR)
  5. Agent Summary sheet (Counselled & Warned agent breakdowns)

Uses Single-Pass Zero-Memory Streaming Engine (core.stream_engine):
- O(1) Memory Footprint (< 35MB RAM)
- Direct XML disk streaming for massive datasets
"""

import sys
import os
import gc
import logging
from datetime import datetime, timedelta
from pathlib import Path
from typing import Union, List, Dict, Any, Optional

import openpyxl
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.formatting.rule import CellIsRule
from openpyxl.utils import get_column_letter

# Ensure server root is in sys.path
SERVER_ROOT = Path(__file__).resolve().parent.parent
if str(SERVER_ROOT) not in sys.path:
    sys.path.insert(0, str(SERVER_ROOT))

try:
    from config.dc_config import ALLOWED_SOURCE_DCS, ALLOWED_DCS_SET
except ImportError:
    from dc_config import ALLOWED_SOURCE_DCS, ALLOWED_DCS_SET

from core.stream_engine import (
    XmlSheetWriter,
    assemble_stream_workbook,
    open_stream_reader,
    get_sheet_names,
    ColumnFinder
)

log = logging.getLogger("ei_stream_server.ei_generator")

ALLOWED_SOURCE_DC = ALLOWED_SOURCE_DCS

IDENTITY_COLS = 3
BLOCK_SIZE    = 6
IDX_OFD       = 0
IDX_FWD_TASK  = 1
IDX_FWD_1K    = 2
IDX_OFP       = 3
IDX_REV_TASK  = 4
IDX_REV_1K    = 5

C_FWD_TITLE   = "1E1B4B"
C_FWD_HDR     = "312E81"
C_REV_TITLE   = "581C87"
C_REV_HDR     = "6B21A8"
C_HDR_FONT    = "FFFFFF"
C_BORDER      = "CBD5E1"

CF_GREEN_BG   = "BBEFCF"
CF_GREEN_FONT = "166534"
CF_YELLOW_BG  = "FEF08A"
CF_YELLOW_FONT= "854D0E"
CF_RED_BG     = "FECACA"
CF_RED_FONT   = "991B1B"

def _fill(hex_color):
    return PatternFill("solid", fgColor=hex_color)

def _font(hex_color, bold=False, size=10):
    return Font(color=hex_color, bold=bold, size=size)

def _border(hex_color=C_BORDER):
    side = Side(style="thin", color=hex_color)
    return Border(left=side, right=side, top=side, bottom=side)

def _center():
    return Alignment(horizontal="center", vertical="center")

def _fmt_date(dt):
    if not isinstance(dt, datetime):
        return str(dt)
    months = ['Jan','Feb','Mar','Apr','May','Jun',
              'Jul','Aug','Sep','Oct','Nov','Dec']
    return f"{dt.day}-{months[dt.month-1]}-{dt.year}"

def _safe_float(val, fallback=0.0):
    try:
        return float(val) if val is not None else fallback
    except (TypeError, ValueError):
        return fallback


def parse_task_per_1k_rows(all_rows: List[List[Any]]):
    if not all_rows:
        raise ValueError("Sheet 'Task_per_1k' is empty.")

    row1 = all_rows[0]
    block_starts = []
    for i in range(IDENTITY_COLS, len(row1)):
        v = row1[i]
        if v is not None and str(v).strip() != '':
            block_starts.append((i, v))

    if not block_starts:
        raise ValueError("No date/WTD blocks found in Task_per_1k row 1.")

    raw_rows = []
    for r_idx in range(2, len(all_rows)):
        row = all_rows[r_idx]
        if not row:
            continue
        dc = row[0] if len(row) > 0 else None
        if dc is None:
            continue
        dc = str(dc).strip()
        if not dc:
            continue
        region = row[1] if len(row) > 1 else ''
        city   = row[2] if len(row) > 2 else ''
        raw_rows.append((row, dc, region, city))

    blocks = []
    for col_start, label in block_starts:
        is_wtd = isinstance(label, str) and label.strip().upper() == 'WTD'
        block_rows = []
        for (row, dc, region, city) in raw_rows:
            if dc not in ALLOWED_SOURCE_DC:
                continue
            ofd      = _safe_float(row[col_start + IDX_OFD] if len(row) > col_start + IDX_OFD else 0)
            fwd_task = _safe_float(row[col_start + IDX_FWD_TASK] if len(row) > col_start + IDX_FWD_TASK else 0)
            fwd_1k   = _safe_float(row[col_start + IDX_FWD_1K] if len(row) > col_start + IDX_FWD_1K else 0)
            ofp      = _safe_float(row[col_start + IDX_OFP] if len(row) > col_start + IDX_OFP else 0)
            rev_task = _safe_float(row[col_start + IDX_REV_TASK] if len(row) > col_start + IDX_REV_TASK else 0)
            rev_1k   = _safe_float(row[col_start + IDX_REV_1K] if len(row) > col_start + IDX_REV_1K else 0)

            if ofd == 0 and fwd_task == 0 and ofp == 0 and rev_task == 0:
                continue

            block_rows.append({
                'dc': dc, 'region': region, 'city': city,
                'ofd': ofd, 'fwd_task': fwd_task, 'fwd_1k': fwd_1k,
                'ofp': ofp, 'rev_task': rev_task, 'rev_1k': rev_1k,
            })

        blocks.append({'label': label, 'is_wtd': is_wtd, 'rows': block_rows})

    return blocks


def select_daily_block(blocks):
    today     = datetime.now().date()
    yesterday = today - timedelta(days=1)
    daily_blocks = [b for b in blocks if not b['is_wtd']]

    for b in daily_blocks:
        lbl = b['label']
        if isinstance(lbl, datetime) and lbl.date() == yesterday:
            return b

    dated = [(b['label'].date(), b) for b in daily_blocks if isinstance(b['label'], datetime)]
    if dated:
        dated.sort(key=lambda x: x[0], reverse=True)
        return dated[0][1]

    if daily_blocks:
        return daily_blocks[0]

    raise ValueError("No daily date blocks found.")


def select_wtd_block(blocks):
    for b in blocks:
        if b['is_wtd']:
            return b
    if blocks:
        return blocks[-1]
    raise ValueError("No WTD block found.")


def build_date_range(blocks):
    dates = [b['label'] for b in blocks if not b['is_wtd'] and isinstance(b['label'], datetime)]
    if not dates:
        return ''
    dates.sort()
    if len(dates) == 1:
        return _fmt_date(dates[0])
    return f"{_fmt_date(dates[0])} - {_fmt_date(dates[-1])}"


def write_summary_sheet(wb, daily_block, wtd_block, date_range_str):
    ws = wb.create_sheet("SUMMARY")
    ws.sheet_view.showGridLines = False

    def _block_date_label(b):
        lbl = b['label']
        return _fmt_date(lbl) if isinstance(lbl, datetime) else str(lbl)

    daily_lbl = _block_date_label(daily_block)
    wtd_lbl   = f"WTD ({date_range_str})" if date_range_str else "WTD"

    def _side_by_side(start_row, title_daily, title_wtd, block_d, block_w):
        headers = [
            "Source_DC", "Region", "City",
            "OFD", "FWD Task", "FWD Task per 1k",
            "OFP", "REV Task", "REV Task per 1k",
        ]

        ws.merge_cells(start_row=start_row, start_column=1, end_row=start_row, end_column=9)
        ws.cell(start_row, 1, title_daily)
        ws.cell(start_row, 1).fill      = _fill(C_FWD_TITLE)
        ws.cell(start_row, 1).font      = _font(C_HDR_FONT, bold=True, size=11)
        ws.cell(start_row, 1).alignment = _center()
        for c in range(1, 10):
            ws.cell(start_row, c).border = _border(C_BORDER)

        ws.cell(start_row, 10, "")

        ws.merge_cells(start_row=start_row, start_column=11, end_row=start_row, end_column=19)
        ws.cell(start_row, 11, title_wtd)
        ws.cell(start_row, 11).fill      = _fill(C_REV_TITLE)
        ws.cell(start_row, 11).font      = _font(C_HDR_FONT, bold=True, size=11)
        ws.cell(start_row, 11).alignment = _center()
        for c in range(11, 20):
            ws.cell(start_row, c).border = _border(C_BORDER)

        r_hdr = start_row + 1
        for i, h in enumerate(headers):
            cell = ws.cell(r_hdr, i + 1, h)
            cell.fill      = _fill(C_FWD_HDR)
            cell.font      = _font(C_HDR_FONT, bold=True)
            cell.alignment = _center()
            cell.border    = _border(C_BORDER)

        for i, h in enumerate(headers):
            cell = ws.cell(r_hdr, i + 11, h)
            cell.fill      = _fill(C_REV_HDR)
            cell.font      = _font(C_HDR_FONT, bold=True)
            cell.alignment = _center()
            cell.border    = _border(C_BORDER)

        rows_d = {r['dc']: r for r in block_d['rows']}
        rows_w = {r['dc']: r for r in block_w['rows']}

        all_dcs = list(ALLOWED_SOURCE_DC)
        all_dcs = [dc for dc in all_dcs if dc in rows_d or dc in rows_w]

        current_row = r_hdr + 1
        for dc in all_dcs:
            rd = rows_d.get(dc)
            rw = rows_w.get(dc)

            if rd:
                vals_d = [
                    rd['dc'], rd['region'], rd['city'],
                    rd['ofd'], rd['fwd_task'], rd['fwd_1k'],
                    rd['ofp'], rd['rev_task'], rd['rev_1k'],
                ]
            else:
                vals_d = [dc, rw['region'] if rw else '', rw['city'] if rw else '', 0, 0, 0, 0, 0, 0]

            if rw:
                vals_w = [
                    rw['dc'], rw['region'], rw['city'],
                    rw['ofd'], rw['fwd_task'], rw['fwd_1k'],
                    rw['ofp'], rw['rev_task'], rw['rev_1k'],
                ]
            else:
                vals_w = [dc, rd['region'] if rd else '', rd['city'] if rd else '', 0, 0, 0, 0, 0, 0]

            for i, v in enumerate(vals_d):
                cell = ws.cell(current_row, i + 1, v)
                cell.alignment = _center()
                cell.border    = _border(C_BORDER)
                cell.font      = _font("000000")
                if i in (3, 4, 6, 7):
                    cell.number_format = '#,##0'
                elif i in (5, 8):
                    cell.number_format = '0.00'

            for i, v in enumerate(vals_w):
                cell = ws.cell(current_row, i + 11, v)
                cell.alignment = _center()
                cell.border    = _border(C_BORDER)
                cell.font      = _font("000000")
                if i in (3, 4, 6, 7):
                    cell.number_format = '#,##0'
                elif i in (5, 8):
                    cell.number_format = '0.00'

            current_row += 1

        r_tot = current_row

        tot_d_ofd      = sum(r['ofd']      for r in rows_d.values())
        tot_d_fwd_task = sum(r['fwd_task'] for r in rows_d.values())
        tot_d_fwd_1k   = (tot_d_fwd_task / tot_d_ofd * 1000) if tot_d_ofd > 0 else 0
        tot_d_ofp      = sum(r['ofp']      for r in rows_d.values())
        tot_d_rev_task = sum(r['rev_task'] for r in rows_d.values())
        tot_d_rev_1k   = (tot_d_rev_task / tot_d_ofp * 1000) if tot_d_ofp > 0 else 0

        tot_w_ofd      = sum(r['ofd']      for r in rows_w.values())
        tot_w_fwd_task = sum(r['fwd_task'] for r in rows_w.values())
        tot_w_fwd_1k   = (tot_w_fwd_task / tot_w_ofd * 1000) if tot_w_ofd > 0 else 0
        tot_w_ofp      = sum(r['ofp']      for r in rows_w.values())
        tot_w_rev_task = sum(r['rev_task'] for r in rows_w.values())
        tot_w_rev_1k   = (tot_w_rev_task / tot_w_ofp * 1000) if tot_w_ofp > 0 else 0

        tot_d_vals = ["Total", "", "", tot_d_ofd, tot_d_fwd_task, tot_d_fwd_1k, tot_d_ofp, tot_d_rev_task, tot_d_rev_1k]
        tot_w_vals = ["Total", "", "", tot_w_ofd, tot_w_fwd_task, tot_w_fwd_1k, tot_w_ofp, tot_w_rev_task, tot_w_rev_1k]

        for i, v in enumerate(tot_d_vals):
            cell = ws.cell(r_tot, i + 1, v)
            cell.alignment = _center()
            cell.font      = _font(C_HDR_FONT, bold=True)
            cell.fill      = _fill(C_FWD_HDR)
            cell.border    = _border(C_BORDER)
            if i in (3, 4, 6, 7):
                cell.number_format = '#,##0'
            elif i in (5, 8):
                cell.number_format = '0.00'

        for i, v in enumerate(tot_w_vals):
            cell = ws.cell(r_tot, i + 11, v)
            cell.alignment = _center()
            cell.font      = _font(C_HDR_FONT, bold=True)
            cell.fill      = _fill(C_REV_HDR)
            cell.border    = _border(C_BORDER)
            if i in (3, 4, 6, 7):
                cell.number_format = '#,##0'
            elif i in (5, 8):
                cell.number_format = '0.00'

        return r_tot

    last_row = _side_by_side(1, daily_lbl, wtd_lbl, daily_block, wtd_block)

    data_start = 3
    data_end   = last_row - 1

    rule_green  = CellIsRule(operator='lessThan', formula=['2.5'],
                             fill=_fill(CF_GREEN_BG), font=_font(CF_GREEN_FONT, bold=True))
    rule_yellow = CellIsRule(operator='between', formula=['2.5', '3.5'],
                             fill=_fill(CF_YELLOW_BG), font=_font(CF_YELLOW_FONT, bold=True))
    rule_red    = CellIsRule(operator='greaterThan', formula=['3.5'],
                             fill=_fill(CF_RED_BG), font=_font(CF_RED_FONT, bold=True))

    for col_letter in ['F', 'I', 'P', 'S']:
        rng = f"{col_letter}{data_start}:{col_letter}{data_end}"
        ws.conditional_formatting.add(rng, rule_green)
        ws.conditional_formatting.add(rng, rule_yellow)
        ws.conditional_formatting.add(rng, rule_red)

    col_widths = {
        'A': 14, 'B': 12, 'C': 16, 'D': 10, 'E': 10, 'F': 14, 'G': 10, 'H': 10, 'I': 14,
        'J': 4,
        'K': 14, 'L': 12, 'M': 16, 'N': 10, 'O': 10, 'P': 14, 'Q': 10, 'R': 10, 'S': 14,
    }
    for col, width in col_widths.items():
        ws.column_dimensions[col].width = width


def write_agent_summary_tab(wb, agent_counts):
    ws = wb.create_sheet("Agent Summary")
    ws.sheet_view.showGridLines = False

    dc_order = ALLOWED_SOURCE_DC
    agent_rows      = []
    counselled_rows = []
    warned_rows     = []

    for dc in dc_order:
        if dc not in agent_counts:
            continue
        for agent in sorted(agent_counts[dc]):
            cnt = agent_counts[dc][agent]
            la  = agent.lower()
            agent_rows.append([dc, agent, cnt])
            if cnt > 2 and la not in ('#n/a', 'n/a'):
                counselled_rows.append([dc, agent, cnt])
            if cnt > 5 and la not in ('#n/a', 'n/a'):
                warned_rows.append([dc, agent, cnt])

    counselled_rows.sort(key=lambda x: x[2], reverse=True)
    warned_rows.sort(key=lambda x: x[2], reverse=True)

    max_rows = max(len(agent_rows), len(counselled_rows), len(warned_rows), 1)

    output = []
    output.append(['Agent Summary', '', '', '', 'Agent to be counselled', '', '', '', 'Agents to be Warned', '', ''])
    output.append(['Source_DC', 'Agent Name', 'Count', '', 'Source_DC', 'Agent Name', 'Count', '', 'Source_DC', 'Agent Name', 'Count'])
    for i in range(max_rows):
        r1 = agent_rows[i]      if i < len(agent_rows)      else ['', '', '']
        r2 = counselled_rows[i] if i < len(counselled_rows) else ['', '', '']
        r3 = warned_rows[i]     if i < len(warned_rows)     else ['', '', '']
        output.append(r1 + [''] + r2 + [''] + r3)

    for r_idx, row in enumerate(output, start=1):
        for c_idx, val in enumerate(row, start=1):
            ws.cell(row=r_idx, column=c_idx, value=val).alignment = _center()

    if len(output) > 1:
        def _title(row, c1, c2, bg):
            ws.merge_cells(start_row=row, start_column=c1, end_row=row, end_column=c2)
            cell = ws.cell(row, c1)
            cell.fill = _fill(bg)
            cell.font = _font(C_HDR_FONT, bold=True, size=11)
            cell.alignment = _center()

        def _subhdr(row, c1, c2, bg):
            for c in range(c1, c2 + 1):
                cell = ws.cell(row, c)
                cell.fill = _fill(bg)
                cell.font = _font(C_HDR_FONT, bold=True)
                cell.alignment = _center()

        _title(1, 1, 3, C_FWD_TITLE)
        _title(1, 5, 7, C_REV_TITLE)
        _title(1, 9, 11, "991B1B")
        _subhdr(2, 1, 3, C_FWD_HDR)
        _subhdr(2, 5, 7, C_REV_HDR)
        _subhdr(2, 9, 11, "B91C1C")

        bd = _border(C_BORDER)
        def _bdr(rs, nr, cs, nc):
            for r in range(rs, rs + nr):
                for c in range(cs, cs + nc):
                    ws.cell(r, c).border = bd

        data_rows = len(output)
        _bdr(1, data_rows, 1, 3)
        _bdr(1, data_rows, 5, 3)
        _bdr(1, data_rows, 9, 3)

    for c, w in [(1,10),(2,28),(3,8),(4,3),(5,10),(6,28),(7,8),(8,3),(9,10),(10,28),(11,8)]:
        ws.column_dimensions[get_column_letter(c)].width = w


def generate_ei_report(source_file_path: Union[str, Path], output_file_path: Union[str, Path]) -> Path:
    source_path = Path(source_file_path)
    output_path = Path(output_file_path)

    log.info(f"Generating EI report (Single-Pass Stream Mode): {source_path} → {output_path}")

    # 1. Parse Task_per_1k tab
    log.info("Phase 1: Parsing Task_per_1k sheet")
    task_rows = []
    with open_stream_reader(source_path, sheet_name='Task_per_1k', header_row=1) as (t_hdr, t_iter):
        task_rows.append(t_hdr)
        for r in t_iter:
            task_rows.append(r)

    blocks = parse_task_per_1k_rows(task_rows)
    daily_block = select_daily_block(blocks)
    wtd_block   = select_wtd_block(blocks)
    date_range  = build_date_range(blocks)
    log.info(f"  Daily block: {daily_block['label']}, WTD block: {wtd_block['label']}, range: {date_range}")

    # 2. Build Summary Workbook
    log.info("Phase 2: Building Summary Sheets")
    wb_out = openpyxl.Workbook()
    wb_out.remove(wb_out.active)
    write_summary_sheet(wb_out, daily_block, wtd_block, date_range)

    # 3. Stream Raw Tab
    all_sheet_names = get_sheet_names(source_path)
    sheet_map = {s.lower(): s for s in all_sheet_names}

    raw_sheet_name = None
    for cand in ['raw', 'raw_data', 'sheet1']:
        if cand in sheet_map:
            raw_sheet_name = sheet_map[cand]
            break

    if not raw_sheet_name and all_sheet_names:
        raw_sheet_name = all_sheet_names[0]

    agent_counts: Dict[str, Dict[str, int]] = {}
    stream_writers: List[XmlSheetWriter] = []

    if raw_sheet_name:
        log.info(f"Phase 3: Streaming Raw tab '{raw_sheet_name}'")
        with open_stream_reader(source_path, sheet_name=raw_sheet_name) as (raw_headers, raw_iter):
            cf = ColumnFinder(raw_headers, {
                'tracking': ['tracking_id', 'tracking id', 'tracking_no', 'trackingno', 'waybill', 'awb', 'shipment'],
                'sdc': ['source_dc', 'source dc', 'dc', 'hub'],
                'fwd_agt': ['fwd_agent_name', 'fwd_delivery_agent', 'delivery_agent', 'agent_name'],
                'rev_agt': ['rev_agent_name', 'rev_delivery_agent', 'pickup_agent', 'agent_name']
            })

            track_idx = cf['tracking']
            dc_idx    = cf['sdc']
            fwd_agt_idx = cf['fwd_agt']
            rev_agt_idx = cf['rev_agt']

            writer_filtered = XmlSheetWriter("Filtered", raw_headers)
            writer_fwd      = XmlSheetWriter("FWD EI", raw_headers)
            writer_rev      = XmlSheetWriter("REVERSE EI", raw_headers)

            stream_writers = [writer_filtered, writer_fwd, writer_rev]

            with writer_filtered, writer_fwd, writer_rev:
                for row in raw_iter:
                    if not row or len(row) <= dc_idx:
                        continue
                    raw_dc = row[dc_idx]
                    if raw_dc is None:
                        continue
                    dc_clean = str(raw_dc).strip()

                    if dc_clean in ALLOWED_SOURCE_DC or dc_clean.upper() in ALLOWED_DCS_SET:
                        writer_filtered.write_row(row)

                        tno = str(row[track_idx] or '').strip().upper() if len(row) > track_idx and row[track_idx] is not None else ''

                        if tno.startswith('MYSC') or tno.startswith('MYSP'):
                            writer_fwd.write_row(row)
                            agent = str(row[fwd_agt_idx] or '').strip() if len(row) > fwd_agt_idx and row[fwd_agt_idx] is not None else ''
                            if not agent: agent = '#N/A'
                            agent_counts.setdefault(dc_clean, {})
                            agent_counts[dc_clean][agent] = agent_counts[dc_clean].get(agent, 0) + 1

                        elif tno.startswith('MYSR'):
                            writer_rev.write_row(row)
                            agent = str(row[rev_agt_idx] or '').strip() if len(row) > rev_agt_idx and row[rev_agt_idx] is not None else ''
                            if not agent: agent = '#N/A'
                            agent_counts.setdefault(dc_clean, {})
                            agent_counts[dc_clean][agent] = agent_counts[dc_clean].get(agent, 0) + 1

    # 4. Write Agent Summary Sheet
    write_agent_summary_tab(wb_out, agent_counts)

    # 5. Assemble final streaming workbook
    log.info("Phase 4: Assembling final workbook")
    assemble_stream_workbook(wb_out, stream_writers, output_path)
    log.info(f"EI Summary Report successfully generated: {output_path.name}")
    return output_path
