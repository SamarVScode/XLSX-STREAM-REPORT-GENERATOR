#!/usr/bin/env python3
"""
NPS Report Generator Module for ei_report_server
================================================
Reads 'Data' sheet from NPS Excel file, filters rows where Source DC is in allowed list,
computes Promoter/Neutral/Detractor stats by DC and by Agent, and generates output workbook:
  1. summary Sheet (Response Breakdown tables by DC & Agent with NPS%)
  2. raw Sheet (Filtered raw dataset)
"""

import sys
import logging
from pathlib import Path
from collections import defaultdict
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# Ensure current directory is in sys.path for dc_config import
CURRENT_DIR = Path(__file__).resolve().parent
if str(CURRENT_DIR) not in sys.path:
    sys.path.insert(0, str(CURRENT_DIR))

try:
    from config.dc_config import ALLOWED_DCS_SET
except ImportError:
    from dc_config import ALLOWED_DCS_SET

log = logging.getLogger("ei_stream_server.nps_report")

def nps_pct(p, n, d):
    total = p + n + d
    if total == 0:
        return 0
    return round((p - d) / total * 100)

def generate_nps_report(input_file: Path, output_file: Path):
    import gc
    log.info(f"Loading input workbook for NPS Report (streaming mode): {input_file}")
    wb_in = openpyxl.load_workbook(str(input_file), data_only=True, read_only=True)
    if 'Data' not in wb_in.sheetnames:
        sheet_name = wb_in.sheetnames[0]
        log.warning(f"Sheet 'Data' not found. Using first sheet: '{sheet_name}'")
        ws_data = wb_in[sheet_name]
    else:
        ws_data = wb_in['Data']

    row_iter = ws_data.iter_rows(values_only=True)
    headers_raw = next(row_iter, None)
    headers = list(headers_raw) if headers_raw else []
    
    # Locate column indices dynamically if available, otherwise use defaults
    source_dc_idx = 28
    option_idx = 4
    agent_idx = 22
    
    for idx, h in enumerate(headers):
        if h is not None:
            h_str = str(h).strip().lower()
            if h_str in ('source_dc', 'source dc'):
                source_dc_idx = idx
            elif h_str in ('option', 'nps_option', 'nps option'):
                option_idx = idx
            elif h_str in ('agent_name', 'agent name', 'agent'):
                agent_idx = idx

    filtered_rows = []
    for row in row_iter:
        if len(row) > source_dc_idx:
            source_dc = str(row[source_dc_idx]).strip().upper() if row[source_dc_idx] is not None else ''
            if source_dc in ALLOWED_DCS_SET:
                filtered_rows.append(row)

    wb_in.close()
    log.info(f"Filtered {len(filtered_rows)} rows matching allowed DCs")

    # Stats aggregation
    dc_stats = defaultdict(lambda: {'P': 0, 'N': 0, 'D': 0})
    agent_stats = defaultdict(lambda: {'P': 0, 'N': 0, 'D': 0})

    for row in filtered_rows:
        option = str(row[option_idx]).strip() if len(row) > option_idx and row[option_idx] is not None else ''
        source_dc = str(row[source_dc_idx]).strip().upper() if len(row) > source_dc_idx and row[source_dc_idx] is not None else ''
        agent = str(row[agent_idx]).strip() if len(row) > agent_idx and row[agent_idx] is not None else ''

        if option == 'Promoter':
            dc_stats[source_dc]['P'] += 1
            agent_stats[(source_dc, agent)]['P'] += 1
        elif option == 'Neutral':
            dc_stats[source_dc]['N'] += 1
            agent_stats[(source_dc, agent)]['N'] += 1
        elif option == 'Detractor':
            dc_stats[source_dc]['D'] += 1
            agent_stats[(source_dc, agent)]['D'] += 1

    wb_out = openpyxl.Workbook()

    # --- Summary sheet ---
    ws_sum = wb_out.active
    ws_sum.title = 'Summary'
    ws_sum.sheet_view.showGridLines = False

    dark_blue_fill = PatternFill(start_color='FF1A365D', end_color='FF1A365D', fill_type='solid')
    med_blue_fill = PatternFill(start_color='FF2B6CB0', end_color='FF2B6CB0', fill_type='solid')
    green_hdr_fill = PatternFill(start_color='FFC6F6D5', end_color='FFC6F6D5', fill_type='solid')
    yellow_hdr_fill = PatternFill(start_color='FFFFFCBF', end_color='FFFFFCBF', fill_type='solid')
    red_hdr_fill = PatternFill(start_color='FFFED7D7', end_color='FFFED7D7', fill_type='solid')
    nps_green_fill = PatternFill(start_color='FFC6F6D5', end_color='FFC6F6D5', fill_type='solid')
    nps_red_fill = PatternFill(start_color='FFFED7D7', end_color='FFFED7D7', fill_type='solid')

    bold_font = Font(bold=True, size=11)
    title_font = Font(bold=True, size=11, color='FFFFFFFF')
    hdr_white_font = Font(bold=True, size=11, color='FFFFFFFF')
    nps_green_font = Font(bold=True, size=11)
    nps_red_font = Font(bold=True, size=11, color='FFC53030')

    thin = Side(style='thin')
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    # Titles
    ws_sum.merge_cells('B1:F1')
    ws_sum['B1'] = 'Response Breakdown'
    ws_sum['B1'].font = title_font
    ws_sum['B1'].fill = dark_blue_fill
    ws_sum['B1'].alignment = Alignment(horizontal='center')

    ws_sum.merge_cells('J1:N1')
    ws_sum['J1'] = 'Response Breakdown'
    ws_sum['J1'].font = title_font
    ws_sum['J1'].fill = dark_blue_fill
    ws_sum['J1'].alignment = Alignment(horizontal='center')

    # Headers
    dc_headers = ['DC', 'P', 'N', 'D', 'Total', 'NPS%']
    agent_headers = ['DC', 'Agent', 'P', 'N', 'D', 'Total', 'NPS%']

    for i, h in enumerate(dc_headers, 1):
        cell = ws_sum.cell(row=2, column=i, value=h)
        cell.font = hdr_white_font
        cell.alignment = Alignment(horizontal='center')
        cell.border = border
        if h == 'P':
            cell.fill = green_hdr_fill
            cell.font = bold_font
        elif h == 'N':
            cell.fill = yellow_hdr_fill
            cell.font = bold_font
        elif h == 'D':
            cell.fill = red_hdr_fill
            cell.font = bold_font
        else:
            cell.fill = med_blue_fill

    for i, h in enumerate(agent_headers, 8):
        cell = ws_sum.cell(row=2, column=i, value=h)
        cell.font = hdr_white_font
        cell.alignment = Alignment(horizontal='center')
        cell.border = border
        if h == 'P':
            cell.fill = green_hdr_fill
            cell.font = bold_font
        elif h == 'N':
            cell.fill = yellow_hdr_fill
            cell.font = bold_font
        elif h == 'D':
            cell.fill = red_hdr_fill
            cell.font = bold_font
        else:
            cell.fill = med_blue_fill

    # DC summary rows
    sorted_dcs = sorted(dc_stats.keys())
    row_num = 3
    for dc in sorted_dcs:
        stats = dc_stats[dc]
        total = stats['P'] + stats['N'] + stats['D']
        nps = nps_pct(stats['P'], stats['N'], stats['D'])
        values = [dc, stats['P'], stats['N'], stats['D'], total, nps]
        for i, v in enumerate(values, 1):
            cell = ws_sum.cell(row=row_num, column=i, value=v)
            cell.alignment = Alignment(horizontal='center')
            cell.border = border
            if i == 6:
                if nps > 85:
                    cell.fill = nps_green_fill
                    cell.font = nps_green_font
                else:
                    cell.fill = nps_red_fill
                    cell.font = nps_red_font
        row_num += 1

    # Agent summary rows (grouped by DC)
    row_num = 3
    for dc in sorted_dcs:
        agents_in_dc = [(k, v) for k, v in agent_stats.items() if k[0] == dc]
        agents_in_dc.sort(key=lambda x: x[0][1])
        for (dc_key, agent), stats in agents_in_dc:
            total = stats['P'] + stats['N'] + stats['D']
            nps = nps_pct(stats['P'], stats['N'], stats['D'])
            values = [dc_key, agent, stats['P'], stats['N'], stats['D'], total, nps]
            for i, v in enumerate(values, 8):
                cell = ws_sum.cell(row=row_num, column=i, value=v)
                cell.alignment = Alignment(horizontal='center')
                cell.border = border
                if i == 14:
                    if nps > 85:
                        cell.fill = nps_green_fill
                        cell.font = nps_green_font
                    else:
                        cell.fill = nps_red_fill
                        cell.font = nps_red_font
            row_num += 1

    # Column widths
    for col in range(1, 7):
        ws_sum.column_dimensions[get_column_letter(col)].width = 12
    ws_sum.column_dimensions['G'].width = 3
    for col in range(8, 15):
        ws_sum.column_dimensions[get_column_letter(col)].width = 18

    # Row heights
    ws_sum.row_dimensions[1].height = 22
    ws_sum.row_dimensions[2].height = 20

    # --- Raw sheet ---
    ws_raw = wb_out.create_sheet('Raw')
    ws_raw.append(headers)
    for row in filtered_rows:
        ws_raw.append(list(row))

    output_path = Path(output_file)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    wb_out.save(str(output_path))
    try:
        wb_out.close()
    except Exception:
        pass

    del wb_out
    gc.collect()
    log.info(f"Successfully generated NPS Report: {output_file}")

def main():
    script_dir = Path(__file__).resolve().parent
    input_file = script_dir.parent / "nps" / "NPS_31-Jul-2026.xlsx"
    output_file = script_dir.parent / "nps" / "output.xlsx"

    if len(sys.argv) > 1:
        input_file = Path(sys.argv[1])
    if len(sys.argv) > 2:
        output_file = Path(sys.argv[2])

    generate_nps_report(input_file, output_file)

if __name__ == "__main__":
    main()
