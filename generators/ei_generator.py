import sys
import os
import logging
from datetime import datetime, timedelta
from pathlib import Path
import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.formatting.rule import CellIsRule
from openpyxl.utils import get_column_letter

from config.dc_config import ALLOWED_SOURCE_DCS, ALLOWED_DCS_SET
from core.stream_engine import stream_sheet_rows, get_sheet_names

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

C_WARN_TITLE = "991B1B"
C_WARN_HDR   = "B91C1C"

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

def parse_task_per_1k_stream(file_path: str):
    rows_iter = stream_sheet_rows(file_path, sheet_name='Task_per_1k')
    try:
        row1 = next(rows_iter)
    except StopIteration:
        raise ValueError("Task_per_1k sheet is empty")

    block_starts = []
    for i in range(IDENTITY_COLS, len(row1)):
        v = row1[i]
        if v is not None and str(v).strip() != '':
            block_starts.append((i, v))

    if not block_starts:
        raise ValueError("No date/WTD blocks found in Task_per_1k row 1.")

    try:
        _ = next(rows_iter) # row 2 header
    except StopIteration:
        pass

    raw_data_rows = []
    for r in rows_iter:
        if not r or len(r) == 0 or r[0] is None:
            continue
        dc = str(r[0]).strip()
        if not dc:
            continue
        region = r[1] if len(r) > 1 else ''
        city = r[2] if len(r) > 2 else ''
        raw_data_rows.append((r, dc, region, city))

    blocks = []
    for col_idx, label in block_starts:
        is_wtd = isinstance(label, str) and label.strip().upper() == 'WTD'
        block_rows = []
        for (r_vals, dc, region, city) in raw_data_rows:
            if dc not in ALLOWED_SOURCE_DC:
                continue

            ofd      = _safe_float(r_vals[col_idx + IDX_OFD] if len(r_vals) > col_idx + IDX_OFD else 0)
            fwd_task = _safe_float(r_vals[col_idx + IDX_FWD_TASK] if len(r_vals) > col_idx + IDX_FWD_TASK else 0)
            fwd_1k   = _safe_float(r_vals[col_idx + IDX_FWD_1K] if len(r_vals) > col_idx + IDX_FWD_1K else 0)
            ofp      = _safe_float(r_vals[col_idx + IDX_OFP] if len(r_vals) > col_idx + IDX_OFP else 0)
            rev_task = _safe_float(r_vals[col_idx + IDX_REV_TASK] if len(r_vals) > col_idx + IDX_REV_TASK else 0)
            rev_1k   = _safe_float(r_vals[col_idx + IDX_REV_1K] if len(r_vals) > col_idx + IDX_REV_1K else 0)

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

def parse_raw_tab_stream(file_path: str):
    rows_iter = stream_sheet_rows(file_path, sheet_name='Raw')
    try:
        raw_headers = next(rows_iter)
    except StopIteration:
        return [], [], {}, None, None, None, None

    headers = [str(h).strip() if h is not None else '' for h in raw_headers]
    col_map = {h.lower().replace(' ', '_'): idx for idx, h in enumerate(headers)}

    def find_col(candidates):
        for c in candidates:
            if c in col_map:
                return col_map[c]
        for h, idx in col_map.items():
            for c in candidates:
                if c in h:
                    return idx
        return None

    dc_idx      = find_col(['source_dc', 'dc', 'sdc', 'source'])
    track_idx   = find_col(['tracking_id', 'tracking_no', 'tracking_number', 'waybill', 'awb', 'task_id'])
    fwd_agt_idx = find_col(['agent_name', 'delivery_agent', 'fwd_agent', 'agent', 'rider_name', 'fe_name'])
    rev_agt_idx = find_col(['pickup_agent', 'rev_agent', 'reverse_agent', 'pickup_fe']) or fwd_agt_idx

    filt_rows = []
    for r in rows_iter:
        if not r or len(r) == 0:
            continue
        dc_val = str(r[dc_idx] or '').strip() if dc_idx is not None and len(r) > dc_idx else ''
        if dc_val in ALLOWED_SOURCE_DC:
            filt_rows.append(r)

    return headers, filt_rows, col_map, track_idx, fwd_agt_idx, rev_agt_idx, dc_idx

def write_summary_sheet(wb, daily_block, wtd_block, date_range_str):
    ws = wb.create_sheet("SUMMARY")
    ws.sheet_view.showGridLines = False

    daily_date_str = _fmt_date(daily_block['label']) if isinstance(daily_block['label'], datetime) else str(daily_block['label'])

    FWD_HEADERS = ['Date', 'Source_DC', 'OFD', 'Forward_Task', 'Fwd_Task_per_1k']
    REV_HEADERS = ['Date', 'Source_DC', 'OFP', 'Reverse_Task', 'Rev_Task_per_1k']

    fwd_daily = sorted(daily_block['rows'], key=lambda r: r['fwd_1k'], reverse=True)
    rev_daily = sorted(daily_block['rows'], key=lambda r: r['rev_1k'], reverse=True)
    fwd_wtd   = sorted(wtd_block['rows'],   key=lambda r: r['fwd_1k'], reverse=True)
    rev_wtd   = sorted(wtd_block['rows'],   key=lambda r: r['rev_1k'], reverse=True)

    max_daily  = max(len(fwd_daily), len(rev_daily))
    max_weekly = max(len(fwd_wtd),   len(rev_wtd))
    max_rows   = max(max_daily, max_weekly)

    output = []
    output.append([
        'Forward EI', '', '', '', '', '',
        'Reverse EI', '', '', '', '', '',
        'Weekly Forward EI', '', '', '', '', '',
        'Weekly Reverse EI', '', '', '', ''
    ])
    output.append(
        FWD_HEADERS + [''] +
        REV_HEADERS + [''] +
        FWD_HEADERS + [''] +
        REV_HEADERS
    )

    def _fwd_daily_row(i):
        if i < len(fwd_daily):
            r = fwd_daily[i]
            return [daily_date_str, r['dc'], r['ofd'], r['fwd_task'], r['fwd_1k']]
        return ['', '', '', '', '']

    def _rev_daily_row(i):
        if i < len(rev_daily):
            r = rev_daily[i]
            return [daily_date_str, r['dc'], r['ofp'], r['rev_task'], r['rev_1k']]
        return ['', '', '', '', '']

    def _fwd_wtd_row(i):
        if i < len(fwd_wtd):
            r = fwd_wtd[i]
            return [date_range_str, r['dc'], r['ofd'], r['fwd_task'], r['fwd_1k']]
        return ['', '', '', '', '']

    def _rev_wtd_row(i):
        if i < len(rev_wtd):
            r = rev_wtd[i]
            return [date_range_str, r['dc'], r['ofp'], r['rev_task'], r['rev_1k']]
        return ['', '', '', '', '']

    for i in range(max_rows):
        output.append(
            _fwd_daily_row(i) + [''] +
            _rev_daily_row(i) + [''] +
            _fwd_wtd_row(i)   + [''] +
            _rev_wtd_row(i)
        )

    for r_idx, row in enumerate(output, start=1):
        for c_idx, val in enumerate(row, start=1):
            cell = ws.cell(row=r_idx, column=c_idx, value=val)
            cell.alignment = _center()

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
            cell.font = _font(C_HDR_FONT, bold=True, size=10)
            cell.alignment = _center()

    _title(1, 1, 5, C_FWD_TITLE)
    _title(1, 7, 11, C_REV_TITLE)
    _title(1, 13, 17, C_FWD_TITLE)
    _title(1, 19, 23, C_REV_TITLE)

    _subhdr(2, 1, 5, C_FWD_HDR)
    _subhdr(2, 7, 11, C_REV_HDR)
    _subhdr(2, 13, 17, C_FWD_HDR)
    _subhdr(2, 19, 23, C_REV_HDR)

    bd = _border(C_BORDER)
    def _apply_borders(r_start, num_rows, c_start, num_cols):
        for r in range(r_start, r_start + num_rows):
            for c in range(c_start, c_start + num_cols):
                ws.cell(r, c).border = bd

    _apply_borders(1, max_rows + 2, 1, 5)
    _apply_borders(1, max_rows + 2, 7, 5)
    _apply_borders(1, max_rows + 2, 13, 5)
    _apply_borders(1, max_rows + 2, 19, 5)

    def _add_cf(col_letter, start_r, end_r):
        rng = f"{col_letter}{start_r}:{col_letter}{end_r}"
        green_fill  = PatternFill(start_color=CF_GREEN_BG,  end_color=CF_GREEN_BG,  fill_type="solid")
        green_font  = Font(color=CF_GREEN_FONT, bold=True)
        yellow_fill = PatternFill(start_color=CF_YELLOW_BG, end_color=CF_YELLOW_BG, fill_type="solid")
        yellow_font = Font(color=CF_YELLOW_FONT, bold=True)
        red_fill    = PatternFill(start_color=CF_RED_BG,    end_color=CF_RED_BG,    fill_type="solid")
        red_font    = Font(color=CF_RED_FONT, bold=True)

        ws.conditional_formatting.add(rng, CellIsRule(operator='lessThan', formula=['10'], fill=green_fill, font=green_font))
        ws.conditional_formatting.add(rng, CellIsRule(operator='between', formula=['10','15'], fill=yellow_fill, font=yellow_font))
        ws.conditional_formatting.add(rng, CellIsRule(operator='greaterThan', formula=['15'], fill=red_fill, font=red_font))

    if max_rows > 0:
        _add_cf('E', 3, max_rows + 2)
        _add_cf('K', 3, max_rows + 2)
        _add_cf('Q', 3, max_rows + 2)
        _add_cf('W', 3, max_rows + 2)

    for c in range(1, 24):
        ws.column_dimensions[get_column_letter(c)].width = 18 if (c % 6 != 0) else 4

def write_filtered_dc_tab(wb, headers, filt_rows):
    ws = wb.create_sheet("Filtered_DC_Data")
    for c, h in enumerate(headers, start=1):
        cell = ws.cell(row=1, column=c, value=h)
        cell.font = Font(bold=True)
        cell.fill = PatternFill("solid", fgColor="E2E8F0")
        cell.alignment = _center()
    for row in filt_rows:
        ws.append(row)
    for c in range(1, len(headers) + 1):
        ws.column_dimensions[get_column_letter(c)].width = 18

def write_fwd_ei_tab(wb, headers, filt_rows, track_idx):
    ws = wb.create_sheet("FORWARD EI")
    fwd_rows = []
    for row in filt_rows:
        tno = str(row[track_idx] or '').strip().upper() if track_idx is not None and len(row) > track_idx else ''
        if tno.startswith('MYSC') or tno.startswith('MYSP'):
            fwd_rows.append(row)
    for c, h in enumerate(headers, start=1):
        cell = ws.cell(row=1, column=c, value=h)
        cell.font = Font(bold=True)
        cell.fill = PatternFill("solid", fgColor="F1F5F9")
        cell.alignment = _center()
    for row in fwd_rows:
        ws.append(row)
    for c in range(1, len(headers) + 1):
        ws.column_dimensions[get_column_letter(c)].width = 18

def write_rev_ei_tab(wb, headers, filt_rows, track_idx):
    ws = wb.create_sheet("REVERSE EI")
    rev_rows = []
    for row in filt_rows:
        tno = str(row[track_idx] or '').strip().upper() if track_idx is not None and len(row) > track_idx else ''
        if tno.startswith('MYSR'):
            rev_rows.append(row)
    for c, h in enumerate(headers, start=1):
        cell = ws.cell(row=1, column=c, value=h)
        cell.font = Font(bold=True)
        cell.fill = PatternFill("solid", fgColor="F1F5F9")
        cell.alignment = _center()
    for row in rev_rows:
        ws.append(row)
    for c in range(1, len(headers) + 1):
        ws.column_dimensions[get_column_letter(c)].width = 18

def write_agent_summary_tab(wb, filt_rows, track_idx, fwd_agt_idx, rev_agt_idx, dc_idx):
    ws = wb.create_sheet("Agent Summary")
    ws.sheet_view.showGridLines = False

    counts = {}
    for row in filt_rows:
        dc  = str(row[dc_idx] or '').strip() if dc_idx is not None and len(row) > dc_idx else ''
        if not dc:
            continue
        tno = str(row[track_idx] or '').strip().upper() if track_idx is not None and len(row) > track_idx else ''
        if tno.startswith('MYSC') or tno.startswith('MYSP'):
            agent = str(row[fwd_agt_idx] or '').strip() if fwd_agt_idx is not None and len(row) > fwd_agt_idx else ''
        elif tno.startswith('MYSR'):
            agent = str(row[rev_agt_idx] or '').strip() if rev_agt_idx is not None and len(row) > rev_agt_idx else ''
        else:
            agent = ''
        if not agent:
            agent = '#N/A'
        counts.setdefault(dc, {})
        counts[dc][agent] = counts[dc].get(agent, 0) + 1

    dc_order = ALLOWED_SOURCE_DC
    agent_rows      = []
    counselled_rows = []
    warned_rows     = []

    for dc in dc_order:
        if dc not in counts:
            continue
        for agent in sorted(counts[dc]):
            cnt = counts[dc][agent]
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
        _title(1, 9, 11, C_WARN_TITLE)
        _subhdr(2, 1, 3, C_FWD_HDR)
        _subhdr(2, 5, 7, C_REV_HDR)
        _subhdr(2, 9, 11, C_WARN_HDR)

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

def generate_ei_report(source_file_path: str, output_file_path: str) -> str:
    log.info(f"Stream generating EI report: {source_file_path} → {output_file_path}")

    sheet_names = get_sheet_names(source_file_path)
    sheet_names_lower = [s.lower() for s in sheet_names]

    if 'task_per_1k' not in sheet_names_lower:
        raise ValueError("Sheet 'Task_per_1k' not found in source workbook")

    log.info("Phase 1: Streaming Task_per_1k tab")
    blocks = parse_task_per_1k_stream(source_file_path)
    daily_block = select_daily_block(blocks)
    wtd_block   = select_wtd_block(blocks)
    date_range  = build_date_range(blocks)

    log.info("Phase 2: Streaming Raw tab")
    if 'raw' not in sheet_names_lower:
        headers, filt_rows, col_map, track_idx, fwd_agt_idx, rev_agt_idx, dc_idx = [], [], {}, None, None, None, None
    else:
        headers, filt_rows, col_map, track_idx, fwd_agt_idx, rev_agt_idx, dc_idx = parse_raw_tab_stream(source_file_path)

    log.info("Phase 3: Building styled output workbook")
    wb_out = openpyxl.Workbook()
    wb_out.remove(wb_out.active)

    write_summary_sheet(wb_out, daily_block, wtd_block, date_range)

    if filt_rows:
        write_filtered_dc_tab(wb_out, headers, filt_rows)
        write_fwd_ei_tab(wb_out, headers, filt_rows, track_idx)
        write_rev_ei_tab(wb_out, headers, filt_rows, track_idx)
        write_agent_summary_tab(wb_out, filt_rows, track_idx, fwd_agt_idx, rev_agt_idx, dc_idx)

    wb_out.save(output_file_path)
    log.info(f"Report generated successfully: {output_file_path}")
    return output_file_path
