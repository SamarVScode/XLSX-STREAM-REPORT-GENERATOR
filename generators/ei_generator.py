import sys
import os
import logging
from pathlib import Path
from datetime import datetime, timedelta, date
from collections import defaultdict
import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.formatting.rule import CellIsRule
from openpyxl.utils import get_column_letter

try:
    from python_calamine import CalamineWorkbook
    HAS_CALAMINE = True
except ImportError:
    HAS_CALAMINE = False

try:
    from config.dc_config import ALLOWED_SOURCE_DCS, ALLOWED_DCS_SET
except ImportError:
    from dc_config import ALLOWED_SOURCE_DCS, ALLOWED_DCS_SET

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
    if dt is None:
        return ""
    if hasattr(dt, 'year') and hasattr(dt, 'month') and hasattr(dt, 'day'):
        months = ['Jan','Feb','Mar','Apr','May','Jun',
                  'Jul','Aug','Sep','Oct','Nov','Dec']
        return f"{dt.day}-{months[dt.month-1]}-{dt.year}"
    if isinstance(dt, str):
        clean = dt.strip().split('T')[0].split(' ')[0]
        for fmt in ("%Y-%m-%d", "%d-%m-%Y", "%d/%m/%Y", "%m/%d/%Y", "%d-%b-%Y"):
            try:
                p = datetime.strptime(clean, fmt)
                months = ['Jan','Feb','Mar','Apr','May','Jun',
                          'Jul','Aug','Sep','Oct','Nov','Dec']
                return f"{p.day}-{months[p.month-1]}-{p.year}"
            except ValueError:
                continue
    return str(dt)

def _get_date_obj(lbl):
    if isinstance(lbl, datetime):
        return lbl.date()
    if hasattr(lbl, 'year') and hasattr(lbl, 'month') and hasattr(lbl, 'day'):
        return lbl
    if isinstance(lbl, str):
        clean = lbl.strip().split('T')[0].split(' ')[0]
        for fmt in ("%Y-%m-%d", "%d-%m-%Y", "%d/%m/%Y", "%m/%d/%Y", "%d-%b-%Y"):
            try:
                return datetime.strptime(clean, fmt).date()
            except ValueError:
                continue
    return None

def _safe_float(val, fallback=0.0):
    try:
        return float(val) if val is not None else fallback
    except (TypeError, ValueError):
        return fallback

def parse_task_per_1k(rows):
    if not rows or len(rows) < 2:
        raise ValueError("No data found in Task_per_1k sheet.")
    row1 = rows[0]

    block_starts = []
    for i in range(IDENTITY_COLS, len(row1)):
        v = row1[i]
        if v is not None:
            block_starts.append((i, v))

    if not block_starts:
        raise ValueError("No date/WTD blocks found in Task_per_1k row 1.")

    raw_rows = []
    for r in range(2, len(rows)):
        row = rows[r]
        if not row:
            continue
        dc = row[0]
        if dc is None:
            continue
        dc = str(dc).strip()
        if not dc:
            continue
        region = row[1] if len(row) > 1 else ""
        city   = row[2] if len(row) > 2 else ""
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
        d = _get_date_obj(b['label'])
        if d == yesterday:
            return b

    dated = [(d, b) for b in daily_blocks if (d := _get_date_obj(b['label'])) is not None]
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
    dates = [_get_date_obj(b['label']) for b in blocks if not b['is_wtd']]
    dates = [d for d in dates if d is not None]
    if not dates:
        return ''
    dates.sort()
    if len(dates) == 1:
        return _fmt_date(dates[0])
    return f"{_fmt_date(dates[0])} - {_fmt_date(dates[-1])}"

def write_summary_sheet(wb, daily_block, wtd_block, date_range_str):
    ws = wb.create_sheet("SUMMARY")
    ws.sheet_view.showGridLines = False

    daily_date_str = _fmt_date(daily_block['label']) if isinstance(daily_block['label'], (datetime, date)) else str(daily_block['label'])

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
    output.append(FWD_HEADERS + [''] + REV_HEADERS + [''] + FWD_HEADERS + [''] + REV_HEADERS)

    for i in range(max_rows):
        f_d = fwd_daily[i] if i < len(fwd_daily) else None
        r_d = rev_daily[i] if i < len(rev_daily) else None
        f_w = fwd_wtd[i]   if i < len(fwd_wtd)   else None
        r_w = rev_wtd[i]   if i < len(rev_wtd)   else None
        output.append([
            daily_date_str if f_d else '', f_d['dc'] if f_d else '',
            f_d['ofd'] if f_d else '', f_d['fwd_task'] if f_d else '',
            round(f_d['fwd_1k'], 2) if f_d else '',
            '',
            daily_date_str if r_d else '', r_d['dc'] if r_d else '',
            r_d['ofp'] if r_d else '', r_d['rev_task'] if r_d else '',
            round(r_d['rev_1k'], 2) if r_d else '',
            '',
            date_range_str if f_w else '', f_w['dc'] if f_w else '',
            f_w['ofd'] if f_w else '', f_w['fwd_task'] if f_w else '',
            round(f_w['fwd_1k'], 2) if f_w else '',
            '',
            date_range_str if r_w else '', r_w['dc'] if r_w else '',
            r_w['ofp'] if r_w else '', r_w['rev_task'] if r_w else '',
            round(r_w['rev_1k'], 2) if r_w else '',
        ])

    for row_idx, row_data in enumerate(output, start=1):
        for col_idx, val in enumerate(row_data, start=1):
            ws.cell(row=row_idx, column=col_idx, value=val)

    total_rows = len(output)

    for r in range(1, total_rows + 1):
        for fmt_col, fmt in [
            (1,'@'),(2,'@'),(3,'0'),(4,'0'),(5,'0.00'),
            (7,'@'),(8,'@'),(9,'0'),(10,'0'),(11,'0.00'),
            (13,'@'),(14,'@'),(15,'0'),(16,'0'),(17,'0.00'),
            (19,'@'),(20,'@'),(21,'0'),(22,'0'),(23,'0.00')
        ]:
            ws.cell(r, fmt_col).number_format = fmt
        for c in range(1, 24):
            ws.cell(r, c).alignment = _center()

    def _style_merged_title(row, c1, c2, bg):
        ws.merge_cells(start_row=row, start_column=c1, end_row=row, end_column=c2)
        cell = ws.cell(row, c1)
        cell.fill = _fill(bg)
        cell.font = _font(C_HDR_FONT, bold=True, size=11)
        cell.alignment = _center()

    def _style_sub_header(row, c1, c2, bg):
        for c in range(c1, c2 + 1):
            cell = ws.cell(row, c)
            cell.fill = _fill(bg)
            cell.font = _font(C_HDR_FONT, bold=True)
            cell.alignment = _center()

    _style_merged_title(1, 1, 5, C_FWD_TITLE)
    _style_merged_title(1, 7, 11, C_REV_TITLE)
    _style_merged_title(1, 13, 17, C_FWD_TITLE)
    _style_merged_title(1, 19, 23, C_REV_TITLE)

    _style_sub_header(2, 1, 5, C_FWD_HDR)
    _style_sub_header(2, 7, 11, C_REV_HDR)
    _style_sub_header(2, 13, 17, C_FWD_HDR)
    _style_sub_header(2, 19, 23, C_REV_HDR)

    bd = _border(C_BORDER)
    def _apply_borders(rs, nr, cs, nc):
        for r in range(rs, rs + nr):
            for c in range(cs, cs + nc):
                ws.cell(r, c).border = bd

    if max_daily > 0:
        _apply_borders(3, max_daily, 1, 5)
        _apply_borders(3, max_daily, 7, 5)
    if max_weekly > 0:
        _apply_borders(3, max_weekly, 13, 5)
        _apply_borders(3, max_weekly, 19, 5)

    rule_fwd_g = CellIsRule(operator='lessThan', formula=['2.5'], fill=_fill(CF_GREEN_BG), font=_font(CF_GREEN_FONT, bold=True))
    rule_fwd_y = CellIsRule(operator='between',  formula=['2.5', '6.0'], fill=_fill(CF_YELLOW_BG), font=_font(CF_YELLOW_FONT, bold=True))
    rule_fwd_r = CellIsRule(operator='greaterThan', formula=['6.0'], fill=_fill(CF_RED_BG), font=_font(CF_RED_FONT, bold=True))

    rule_rev_g = CellIsRule(operator='lessThan', formula=['6.1'], fill=_fill(CF_GREEN_BG), font=_font(CF_GREEN_FONT, bold=True))
    rule_rev_y = CellIsRule(operator='between',  formula=['6.1', '10.0'], fill=_fill(CF_YELLOW_BG), font=_font(CF_YELLOW_FONT, bold=True))
    rule_rev_r = CellIsRule(operator='greaterThan', formula=['10.0'], fill=_fill(CF_RED_BG), font=_font(CF_RED_FONT, bold=True))

    if max_daily > 0:
        daily_end_row = 2 + max_daily
        ws.conditional_formatting.add(f"E3:E{daily_end_row}", rule_fwd_g)
        ws.conditional_formatting.add(f"E3:E{daily_end_row}", rule_fwd_y)
        ws.conditional_formatting.add(f"E3:E{daily_end_row}", rule_fwd_r)
        ws.conditional_formatting.add(f"K3:K{daily_end_row}", rule_rev_g)
        ws.conditional_formatting.add(f"K3:K{daily_end_row}", rule_rev_y)
        ws.conditional_formatting.add(f"K3:K{daily_end_row}", rule_rev_r)

    if max_weekly > 0:
        weekly_end_row = 2 + max_weekly
        ws.conditional_formatting.add(f"Q3:Q{weekly_end_row}", rule_fwd_g)
        ws.conditional_formatting.add(f"Q3:Q{weekly_end_row}", rule_fwd_y)
        ws.conditional_formatting.add(f"Q3:Q{weekly_end_row}", rule_fwd_r)
        ws.conditional_formatting.add(f"W3:W{weekly_end_row}", rule_rev_g)
        ws.conditional_formatting.add(f"W3:W{weekly_end_row}", rule_rev_y)
        ws.conditional_formatting.add(f"W3:W{weekly_end_row}", rule_rev_r)

    col_widths = {
        1:14, 2:12, 3:10, 4:14, 5:16, 6:3,
        7:14, 8:12, 9:10, 10:14, 11:16, 12:3,
        13:16, 14:12, 15:10, 16:14, 17:16, 18:3,
        19:16, 20:12, 21:10, 22:14, 23:16
    }
    for col_idx, width in col_widths.items():
        ws.column_dimensions[get_column_letter(col_idx)].width = width

def parse_raw_rows(rows):
    if not rows:
        return [], [], {}, None, None, None, None
    headers = list(rows[0])
    col_map = {str(h).strip().lower(): i for i, h in enumerate(headers) if h is not None}

    dc_idx = None
    for candidate in ['source_dc', 'source dc', 'dc']:
        if candidate in col_map:
            dc_idx = col_map[candidate]
            break

    track_idx = None
    for candidate in ['final_tracking_no', 'tracking_id', 'tracking id', 'tracking_no', 'tracking no', 'waybill', 'awb', 'task_id']:
        if candidate in col_map:
            track_idx = col_map[candidate]
            break

    fwd_agt_idx = col_map.get('fwd_agent name') or col_map.get('fwd agent') or col_map.get('fwd_agent') or col_map.get('agent_name') or col_map.get('agent') or col_map.get('delivery_agent')
    rev_agt_idx = col_map.get('rev_agent name') or col_map.get('rev agent') or col_map.get('rev_agent') or col_map.get('pickup_agent') or fwd_agt_idx

    if dc_idx is None:
        raise ValueError('Column Source_DC not found in Raw tab.')
    if track_idx is None:
        raise ValueError('Column Final_tracking_no / Tracking_ID not found in Raw tab.')

    filt_data_rows = []
    for row in rows[1:]:
        if not any(v for v in row):
            continue
        dc = str(row[dc_idx] or '').strip()
        if dc in ALLOWED_SOURCE_DC:
            filt_data_rows.append(list(row))

    return headers, filt_data_rows, col_map, track_idx, fwd_agt_idx, rev_agt_idx, dc_idx

def write_filtered_dc_tab(wb, headers, filt_data_rows):
    ws = wb.create_sheet('Filtered_Source_DC')
    ws.append(headers)
    for r in filt_data_rows:
        ws.append(r)
    for col_idx in range(1, len(headers) + 1):
        cell = ws.cell(1, col_idx)
        cell.fill = _fill(C_FWD_HDR)
        cell.font = _font(C_HDR_FONT, bold=True)
        cell.alignment = _center()
        ws.column_dimensions[get_column_letter(col_idx)].width = 16

def write_fwd_ei_tab(wb, headers, filt_data_rows, track_idx):
    ws = wb.create_sheet('FWD EI')
    ws.append(headers)
    for r in filt_data_rows:
        tn = str(r[track_idx] or '')
        if tn.startswith(('MYSC', 'MYSD')):
            ws.append(r)
    for col_idx in range(1, len(headers) + 1):
        cell = ws.cell(1, col_idx)
        cell.fill = _fill(C_FWD_HDR)
        cell.font = _font(C_HDR_FONT, bold=True)
        cell.alignment = _center()
        ws.column_dimensions[get_column_letter(col_idx)].width = 16

def write_rev_ei_tab(wb, headers, filt_data_rows, track_idx):
    ws = wb.create_sheet('REVERSE EI')
    ws.append(headers)
    for r in filt_data_rows:
        tn = str(r[track_idx] or '')
        if tn.startswith(('MYSR', 'MYSP')):
            ws.append(r)
    for col_idx in range(1, len(headers) + 1):
        cell = ws.cell(1, col_idx)
        cell.fill = _fill(C_REV_HDR)
        cell.font = _font(C_HDR_FONT, bold=True)
        cell.alignment = _center()
        ws.column_dimensions[get_column_letter(col_idx)].width = 16

def write_agent_summary_tab(wb, filt_data_rows, track_idx, fwd_agt_idx, rev_agt_idx, dc_idx):
    ws = wb.create_sheet('Agent Summary')
    ws.sheet_view.showGridLines = False

    fwd_counts = defaultdict(lambda: defaultdict(int))
    rev_counts = defaultdict(lambda: defaultdict(int))
    warn_counts = defaultdict(lambda: defaultdict(int))

    for r in filt_data_rows:
        dc = str(r[dc_idx] or '').strip()
        tn = str(r[track_idx] or '')
        fwd_agt = str(r[fwd_agt_idx] or '').strip() if fwd_agt_idx is not None and len(r) > fwd_agt_idx else ''
        rev_agt = str(r[rev_agt_idx] or '').strip() if rev_agt_idx is not None and len(r) > rev_agt_idx else ''

        if tn.startswith(('MYSC', 'MYSD')):
            if fwd_agt: fwd_counts[dc][fwd_agt] += 1
            if rev_agt: warn_counts[dc][rev_agt] += 1
        elif tn.startswith(('MYSR', 'MYSP')):
            if rev_agt: rev_counts[dc][rev_agt] += 1
            if fwd_agt: warn_counts[dc][fwd_agt] += 1

    def _sorted_items(counts_dict):
        items = []
        for dc in sorted(counts_dict.keys()):
            for agt, cnt in sorted(counts_dict[dc].items(), key=lambda x: x[1], reverse=True):
                items.append((dc, agt, cnt))
        return items

    fwd_items  = _sorted_items(fwd_counts)
    rev_items  = _sorted_items(rev_counts)
    warn_items = _sorted_items(warn_counts)
    max_len = max(len(fwd_items), len(rev_items), len(warn_items))

    output = [
        ['Forward Agent Summary', '', '', '', 'Reverse Agent Summary', '', '', '', 'Agent Warning Summary', '', ''],
        ['Source_DC', 'Agent Name', 'Count', '', 'Source_DC', 'Agent Name', 'Count', '', 'Source_DC', 'Agent Name', 'Count']
    ]

    for i in range(max_len):
        f = fwd_items[i]  if i < len(fwd_items)  else ('', '', '')
        r = rev_items[i]  if i < len(rev_items)  else ('', '', '')
        w = warn_items[i] if i < len(warn_items) else ('', '', '')
        output.append([f[0], f[1], f[2] if f[2] else '', '',
                       r[0], r[1], r[2] if r[2] else '', '',
                       w[0], w[1], w[2] if w[2] else ''])

    for row_idx, row_data in enumerate(output, start=1):
        for col_idx, val in enumerate(row_data, start=1):
            ws.cell(row=row_idx, column=col_idx, value=val)

    left_align = Alignment(horizontal="left", vertical="center")
    center_align = Alignment(horizontal="center", vertical="center")

    for r in range(3, len(output) + 1):
        # DC and Counts center aligned
        for c in [1, 3, 5, 7, 9, 11]:
            ws.cell(r, c).alignment = center_align
            if c in [3, 7, 11] and ws.cell(r, c).value:
                ws.cell(r, c).number_format = '#,##0'
        # Agent names left aligned
        for c in [2, 6, 10]:
            ws.cell(r, c).alignment = left_align

    C_WARN_TITLE = "7C2D12"
    C_WARN_HDR   = "9A3412"

    def _title(r, c1, c2, bg):
        ws.merge_cells(start_row=r, start_column=c1, end_row=r, end_column=c2)
        cell = ws.cell(r, c1)
        cell.fill = _fill(bg)
        cell.font = _font(C_HDR_FONT, bold=True, size=11)
        cell.alignment = center_align

    def _subhdr(r, c1, c2, bg):
        for c in range(c1, c2 + 1):
            cell = ws.cell(r, c)
            cell.fill = _fill(bg)
            cell.font = _font(C_HDR_FONT, bold=True)
            cell.alignment = center_align

    _title(1, 1, 3, C_FWD_TITLE)
    _title(1, 5, 7, C_REV_TITLE)
    _title(1, 9, 11, C_WARN_TITLE)
    _subhdr(2, 1, 3, C_FWD_HDR)
    _subhdr(2, 5, 7, C_REV_HDR)
    _subhdr(2, 9, 11, C_WARN_HDR)

    bd = _border(C_BORDER)
    # Apply borders cleanly strictly to existing table rows
    for r in range(1, len(fwd_items) + 3):
        for c in [1, 2, 3]:
            ws.cell(r, c).border = bd

    for r in range(1, len(rev_items) + 3):
        for c in [5, 6, 7]:
            ws.cell(r, c).border = bd

    for r in range(1, len(warn_items) + 3):
        for c in [9, 10, 11]:
            ws.cell(r, c).border = bd

    for c, w in [(1,12),(2,30),(3,10),(4,4),(5,12),(6,30),(7,10),(8,4),(9,12),(10,30),(11,10)]:
        ws.column_dimensions[get_column_letter(c)].width = w

def generate_ei_report(source_file_path: str, output_file_path: str) -> str:
    path = Path(source_file_path)
    log.info(f"Generating EI report: {path} → {output_file_path}")

    log.info("Phase 1: Loading source workbook (Rust zero-copy streaming)")
    task_rows = []
    raw_rows = []

    if HAS_CALAMINE:
        try:
            cal = CalamineWorkbook.from_path(str(path))
            sheet_map = {s.lower(): s for s in cal.sheet_names}
            
            task_sheet = sheet_map.get('task_per_1k', cal.sheet_names[0])
            task_rows = cal.get_sheet_by_name(task_sheet).to_python()
            
            if 'raw' in sheet_map:
                raw_sheet = cal.get_sheet_by_name(sheet_map['raw'])
                raw_iter = iter(raw_sheet.iter_rows())
                h_row = next(raw_iter, None)
                if h_row:
                    raw_rows = [list(h_row)]
                    dc_cand_idx = None
                    for idx, h in enumerate(h_row):
                        if str(h or '').strip().lower() in ('source_dc', 'source dc', 'dc'):
                            dc_cand_idx = idx
                            break
                    for r in raw_iter:
                        if dc_cand_idx is not None and len(r) > dc_cand_idx:
                            dc_val = str(r[dc_cand_idx] or '').strip()
                            if dc_val in ALLOWED_SOURCE_DC:
                                raw_rows.append(list(r))
                        else:
                            raw_rows.append(list(r))
        except Exception as ex:
            log.warning(f"Calamine read warning: {ex}. Falling back to openpyxl read_only=True")
            task_rows = []

    if not task_rows:
        wb_src = openpyxl.load_workbook(str(path), data_only=True, read_only=True)
        task_ws = wb_src['Task_per_1k'] if 'Task_per_1k' in wb_src.sheetnames else wb_src.active
        task_rows = [list(r) for r in task_ws.iter_rows(values_only=True)]
        if 'Raw' in wb_src.sheetnames:
            raw_iter = wb_src['Raw'].iter_rows(values_only=True)
            h_row = next(raw_iter, None)
            if h_row:
                raw_rows = [list(h_row)]
                dc_cand_idx = None
                for idx, h in enumerate(h_row):
                    if str(h or '').strip().lower() in ('source_dc', 'source dc', 'dc'):
                        dc_cand_idx = idx
                        break
                for r in raw_iter:
                    if dc_cand_idx is not None and len(r) > dc_cand_idx:
                        dc_val = str(r[dc_cand_idx] or '').strip()
                        if dc_val in ALLOWED_SOURCE_DC:
                            raw_rows.append(list(r))
                    else:
                        raw_rows.append(list(r))
        wb_src.close()

    log.info("Phase 2: Parsing Task_per_1k sheet")
    blocks = parse_task_per_1k(task_rows)
    log.info(f"  Found {len(blocks)} blocks")

    daily_block = select_daily_block(blocks)
    wtd_block   = select_wtd_block(blocks)
    date_range  = build_date_range(blocks)
    log.info(f"  Daily block: {daily_block['label']}, WTD block: {wtd_block['label']}, range: {date_range}")

    log.info("Phase 3: Parsing Raw sheet")
    if not raw_rows:
        headers, filt_rows, col_map, track_idx, fwd_agt_idx, rev_agt_idx, dc_idx = [], [], {}, None, None, None, None
        log.warning("  Raw sheet not found — skipping FWD/REV tabs")
    else:
        headers, filt_rows, col_map, track_idx, fwd_agt_idx, rev_agt_idx, dc_idx = parse_raw_rows(raw_rows)
        log.info(f"  Total filtered rows: {len(filt_rows)}")

    log.info("Phase 4: Writing output workbook")
    wb_out = openpyxl.Workbook()
    wb_out.remove(wb_out.active)

    write_summary_sheet(wb_out, daily_block, wtd_block, date_range)
    log.info("  Summary sheet written")

    if filt_rows:
        write_filtered_dc_tab(wb_out, headers, filt_rows)
        write_fwd_ei_tab(wb_out, headers, filt_rows, track_idx)
        write_rev_ei_tab(wb_out, headers, filt_rows, track_idx)
        write_agent_summary_tab(wb_out, filt_rows, track_idx, fwd_agt_idx, rev_agt_idx, dc_idx)

    log.info("Phase 5: Saving output file")
    wb_out.save(output_file_path)
    try:
        wb_out.close()
    except Exception:
        pass
    del wb_out, raw_rows, filt_rows
    import gc
    gc.collect()
    log.info(f"Report saved: {output_file_path}")
    return str(output_file_path)
