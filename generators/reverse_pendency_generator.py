#!/usr/bin/env python3
"""
Reverse Pendency Report Generator Module for ei_report_server
=============================================================
Reads 'Raw' sheet from input Excel file, filters rows where Region == 'North'
and Source DC is in allowed list, computes Age_Bucket, and generates output workbook:
  1. Summary Sheet (Aging wise report)
  2. P0 reverse pendency Sheet (Aging >= 2 tracking details)
  3. Raw Sheet (Full filtered dataset)
"""

import logging
from pathlib import Path
from collections import defaultdict
from python_calamine import CalamineWorkbook
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter

try:
    from config.dc_config import ALLOWED_SOURCE_DCS, ALLOWED_DCS_SET
except ImportError:
    from dc_config import ALLOWED_SOURCE_DCS, ALLOWED_DCS_SET
AGING_CATEGORIES = ['0-2 Days', '3-5 Days', '6-10 Days', '>10 Days']

log = logging.getLogger("ei_stream_server.reverse_pendency")

def compute_age_bucket(val) -> str:
    if val is None or str(val).strip() == '':
        return '0-2 Days'
    try:
        aging = float(val)
        if aging <= 2:
            return '0-2 Days'
        elif aging <= 5:
            return '3-5 Days'
        elif aging <= 10:
            return '6-10 Days'
        else:
            return '>10 Days'
    except (ValueError, TypeError):
        return '0-2 Days'

def build_summary_sheet(out_wb, filtered_rows, src_dc_idx, age_bucket_idx):
    ws = out_wb.create_sheet(title="Summary", index=0)
    ws.sheet_view.showGridLines = False

    title_fill = PatternFill("solid", fgColor="1E1B4B")
    title_font = Font(name="Calibri", size=12, bold=True, color="FFFFFF")
    header_fill = PatternFill("solid", fgColor="312E81")
    header_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    total_fill = PatternFill("solid", fgColor="E0E7FF")
    total_font = Font(name="Calibri", size=11, bold=True, color="1E1B4B")
    data_font = Font(name="Calibri", size=11, color="1F2937")
    red_fill = PatternFill("solid", fgColor="FECACA")
    red_font = Font(name="Calibri", size=11, bold=True, color="991B1B")
    thin_side = Side(style="thin", color="CBD5E1")
    border = Border(left=thin_side, right=thin_side, top=thin_side, bottom=thin_side)
    center = Alignment(horizontal="center", vertical="center")
    left = Alignment(horizontal="left", vertical="center")

    pivot = defaultdict(lambda: defaultdict(int))
    for row in filtered_rows[1:]:
        dc = str(row[src_dc_idx]).strip().upper() if row[src_dc_idx] else ''
        bucket = str(row[age_bucket_idx]).strip() if row[age_bucket_idx] else '0-2 Days'
        pivot[dc][bucket] += 1

    headers = ["Source DC"] + AGING_CATEGORIES + ["Total Pendency"]
    data = []
    tot_buckets = defaultdict(int)

    for dc in ALLOWED_SOURCE_DCS:
        row_vals = [dc]
        row_tot = 0
        for cat in AGING_CATEGORIES:
            cnt = pivot[dc][cat]
            row_vals.append(cnt)
            row_tot += cnt
            tot_buckets[cat] += cnt
        row_vals.append(row_tot)
        data.append(row_vals)

    totals = ["Total"] + [tot_buckets[cat] for cat in AGING_CATEGORIES] + [sum(tot_buckets.values())]
    data.append(totals)

    end_col = len(headers)
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=end_col)
    title_cell = ws.cell(row=1, column=1, value="Aging wise report")
    title_cell.font = title_font
    title_cell.alignment = center
    for c in range(1, end_col + 1):
        cell = ws.cell(row=1, column=c)
        cell.fill = title_fill
        cell.border = border

    for i, h in enumerate(headers):
        cell = ws.cell(row=2, column=i + 1, value=h)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = center
        cell.border = border

    for r_idx, row_vals in enumerate(data):
        row_num = r_idx + 3
        is_total = row_vals[0] == "Total"
        for c_idx, val in enumerate(row_vals):
            cell = ws.cell(row=row_num, column=c_idx + 1, value=val)
            cell.border = border
            cell.alignment = left if c_idx == 0 else center

            if isinstance(val, (int, float)) and val > 0 and headers[c_idx] in ['3-5 Days', '6-10 Days', '>10 Days']:
                cell.fill = red_fill
                cell.font = red_font
            elif is_total:
                cell.fill = total_fill
                cell.font = total_font
            else:
                cell.font = data_font

    for col in range(1, end_col + 1):
        letter = get_column_letter(col)
        max_len = max(len(str(ws.cell(r, col).value or '')) for r in range(1, len(data) + 3))
        ws.column_dimensions[letter].width = max(max_len + 3, 14)

def build_p0_sheet(out_wb, filtered_rows, header):
    ws = out_wb.create_sheet(title="Critical P0")
    ws.sheet_view.showGridLines = False

    title_fill = PatternFill("solid", fgColor="1E1B4B")
    title_font = Font(name="Calibri", size=12, bold=True, color="FFFFFF")
    header_fill = PatternFill("solid", fgColor="312E81")
    header_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    data_font = Font(name="Calibri", size=11, color="1F2937")
    thin_side = Side(style="thin", color="CBD5E1")
    border = Border(left=thin_side, right=thin_side, top=thin_side, bottom=thin_side)
    center = Alignment(horizontal="center", vertical="center")
    left = Alignment(horizontal="left", vertical="center")

    col_map = {str(h).strip().lower(): idx for idx, h in enumerate(header) if h is not None}
    tn_idx = 0
    for c in ['tracking_number', 'tracking_no', 'tracking_id', 'tracking id', 'waybill', 'awb', 'shipment']:
        if c in col_map:
            tn_idx = col_map[c]
            break

    src_dc_idx = col_map.get('source dc', col_map.get('source_dc', col_map.get('dc', 0)))
    aging_idx = col_map.get('aging', col_map.get('age', 1))
    age_bucket_idx = col_map.get('age_bucket', col_map.get('age bucket', len(header) - 1))
    attempt_idx = col_map.get('attempt_status', col_map.get('status', col_map.get('attempt', 0)))

    p0_rows = []
    for row in filtered_rows[1:]:
        try:
            aging = float(row[aging_idx]) if row[aging_idx] else 0
        except (ValueError, TypeError):
            aging = 0
        if aging >= 2:
            p0_rows.append(row)

    out_headers = ['tracking_number', 'Source DC', 'Aging', 'Age_Bucket', 'Attempt_Status']

    end_col = len(out_headers)
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=end_col)
    title_cell = ws.cell(row=1, column=1, value="P0 reverse pendency (Aging ≥ 2)")
    title_cell.font = title_font
    title_cell.alignment = center
    for c in range(1, end_col + 1):
        cell = ws.cell(row=1, column=c)
        cell.fill = title_fill
        cell.border = border

    for i, h in enumerate(out_headers):
        cell = ws.cell(row=2, column=i + 1, value=h)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = center
        cell.border = border

    for r_idx, row in enumerate(p0_rows):
        row_num = r_idx + 3
        vals = [
            row[tn_idx],
            str(row[src_dc_idx]).strip().upper() if row[src_dc_idx] else '',
            row[aging_idx],
            str(row[age_bucket_idx]).strip() if row[age_bucket_idx] else '',
            str(row[attempt_idx]).strip() if row[attempt_idx] else ''
        ]
        for c_idx, val in enumerate(vals):
            cell = ws.cell(row=row_num, column=c_idx + 1, value=val)
            cell.border = border
            cell.font = data_font
            cell.alignment = left if c_idx in [0, 4] else center

    col_widths = [20, 14, 10, 14, 22]
    for i, w in enumerate(col_widths):
        ws.column_dimensions[get_column_letter(i + 1)].width = w

    return len(p0_rows)

def generate_reverse_pendency_report(input_file: Path, output_file: Path):
    import io
    import zipfile
    import math
    import tempfile
    import gc

    log.info(f"Reading workbook for Reverse Pendency (Zero-Memory Stream Mode): {input_file}")
    wb = openpyxl.load_workbook(str(input_file), read_only=True, data_only=True)

    target_sheet = wb.sheetnames[0]
    for cand in ['raw', 'raw_data', 'data']:
        for s in wb.sheetnames:
            if s.lower() == cand:
                target_sheet = s
                break

    ws = wb[target_sheet]
    row_iter = ws.iter_rows(values_only=True)
    header = next(row_iter, None)
    if not header:
        wb.close()
        raise ValueError(f"Sheet '{target_sheet}' is empty.")

    col_map = {str(h).strip().lower(): idx for idx, h in enumerate(header) if h is not None}
    
    src_dc_idx = 0
    for c in ['source dc', 'source_dc', 'dc']:
        if c in col_map:
            src_dc_idx = col_map[c]
            break

    region_idx = col_map.get('region', None)
    
    aging_idx = 1
    for c in ['aging', 'aging bucket', 'age', 'age_bucket', 'ageing']:
        if c in col_map:
            aging_idx = col_map[c]
            break

    tn_idx = 0
    for c in ['tracking_number', 'tracking_no', 'tracking_id', 'tracking id', 'waybill', 'awb', 'shipment']:
        if c in col_map:
            tn_idx = col_map[c]
            break

    attempt_idx = col_map.get('attempt_status', col_map.get('status', col_map.get('attempt', 0)))

    def esc(val):
        if val is None:
            return ''
        s = str(val)
        return s.replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;').replace('"', '&quot;')

    out_path = Path(output_file)
    out_path.parent.mkdir(parents=True, exist_ok=True)
    temp_raw_xml = Path(tempfile.gettempdir()) / f"temp_raw_{out_path.stem}.xml"
    temp_p0_xml  = Path(tempfile.gettempdir()) / f"temp_p0_{out_path.stem}.xml"

    pivot = defaultdict(lambda: defaultdict(int))
    total_filtered = 0
    p0_count = 0

    raw_headers = list(header) + ['Age_Bucket']
    p0_headers = ['tracking_number', 'Source DC', 'Aging', 'Age_Bucket', 'Attempt_Status']

    f_raw = open(temp_raw_xml, 'wb')
    f_p0  = open(temp_p0_xml, 'wb')

    f_raw.write(b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData>')
    f_p0.write(b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData>')

    # Raw Header
    hdr_xml = ['<row r="1">']
    for c_i, h_val in enumerate(raw_headers, 1):
        hdr_xml.append(f'<c r="{get_column_letter(c_i)}1" t="inlineStr"><is><t>{esc(h_val)}</t></is></c>')
    hdr_xml.append('</row>')
    f_raw.write(''.join(hdr_xml).encode('utf-8'))

    # P0 Title & Header
    p0_title_xml = ['<row r="1">']
    for c_i in range(1, len(p0_headers) + 1):
        val = 'P0 reverse pendency (Aging \u2265 2)' if c_i == 1 else ''
        p0_title_xml.append(f'<c r="{get_column_letter(c_i)}1" t="inlineStr"><is><t>{esc(val)}</t></is></c>')
    p0_title_xml.append('</row><row r="2">')
    for c_i, h_val in enumerate(p0_headers, 1):
        p0_title_xml.append(f'<c r="{get_column_letter(c_i)}2" t="inlineStr"><is><t>{esc(h_val)}</t></is></c>')
    p0_title_xml.append('</row>')
    f_p0.write(''.join(p0_title_xml).encode('utf-8'))

    raw_row_num = 2
    p0_row_num = 3
    raw_chunk = []
    p0_chunk = []
    processed_count = 0

    for row in row_iter:
        processed_count += 1
        if processed_count % 1000 == 0:
            import time
            time.sleep(0.002)

        region = str(row[region_idx]).strip() if (region_idx is not None and len(row) > region_idx and row[region_idx]) else 'North'
        src_dc = str(row[src_dc_idx]).strip().upper() if (len(row) > src_dc_idx and row[src_dc_idx]) else ''
        if (region == 'North' or region_idx is None) and src_dc in ALLOWED_DCS_SET:
            total_filtered += 1
            aging_val = row[aging_idx] if len(row) > aging_idx else 0
            bucket = compute_age_bucket(aging_val)
            pivot[src_dc][bucket] += 1
            
            try:
                aging_f = float(aging_val) if aging_val else 0.0
            except (ValueError, TypeError):
                aging_f = 0.0
            
            # Write to Raw chunk
            r_xml = [f'<row r="{raw_row_num}">']
            for c_i, val in enumerate(row, 1):
                col_let = get_column_letter(c_i)
                if isinstance(val, (int, float)) and not math.isnan(val) and not math.isinf(val):
                    r_xml.append(f'<c r="{col_let}{raw_row_num}"><v>{val}</v></c>')
                else:
                    r_xml.append(f'<c r="{col_let}{raw_row_num}" t="inlineStr"><is><t>{esc(val)}</t></is></c>')
            r_xml.append(f'<c r="{get_column_letter(len(row)+1)}{raw_row_num}" t="inlineStr"><is><t>{bucket}</t></is></c></row>')
            raw_chunk.append(''.join(r_xml))
            raw_row_num += 1
            
            # Write to P0 chunk if aging >= 2
            if aging_f >= 2.0:
                p0_count += 1
                p_vals = [row[tn_idx], src_dc, aging_val, bucket, str(row[attempt_idx] if len(row) > attempt_idx and row[attempt_idx] else '')]
                p_xml = [f'<row r="{p0_row_num}">']
                for c_i, val in enumerate(p_vals, 1):
                    col_let = get_column_letter(c_i)
                    if isinstance(val, (int, float)) and not math.isnan(val) and not math.isinf(val):
                        p_xml.append(f'<c r="{col_let}{p0_row_num}"><v>{val}</v></c>')
                    else:
                        p_xml.append(f'<c r="{col_let}{p0_row_num}" t="inlineStr"><is><t>{esc(val)}</t></is></c>')
                p_xml.append('</row>')
                p0_chunk.append(''.join(p_xml))
                p0_row_num += 1

            if len(raw_chunk) >= 1000:
                f_raw.write(''.join(raw_chunk).encode('utf-8'))
                raw_chunk.clear()
            if len(p0_chunk) >= 1000:
                f_p0.write(''.join(p0_chunk).encode('utf-8'))
                p0_chunk.clear()

    if raw_chunk:
        f_raw.write(''.join(raw_chunk).encode('utf-8'))
    if p0_chunk:
        f_p0.write(''.join(p0_chunk).encode('utf-8'))

    f_raw.write(b'</sheetData></worksheet>')
    f_p0.write(b'</sheetData></worksheet>')
    f_raw.close()
    f_p0.close()
    wb.close()

    # Build Summary sheet in tiny openpyxl workbook (~70 rows, ~50 KB RAM)
    out_wb = Workbook()
    ws_sum = out_wb.active
    ws_sum.title = 'Summary'
    ws_sum.sheet_view.showGridLines = False

    title_fill = PatternFill('solid', fgColor='1E1B4B')
    title_font = Font(name='Calibri', size=12, bold=True, color='FFFFFF')
    header_fill = PatternFill('solid', fgColor='312E81')
    header_font = Font(name='Calibri', size=11, bold=True, color='FFFFFF')
    total_fill = PatternFill('solid', fgColor='E0E7FF')
    total_font = Font(name='Calibri', size=11, bold=True, color='1E1B4B')
    data_font = Font(name='Calibri', size=11, color='1F2937')
    red_fill = PatternFill('solid', fgColor='FECACA')
    red_font = Font(name='Calibri', size=11, bold=True, color='991B1B')
    thin_side = Side(style='thin', color='CBD5E1')
    border = Border(left=thin_side, right=thin_side, top=thin_side, bottom=thin_side)
    center = Alignment(horizontal='center', vertical='center')
    left = Alignment(horizontal='left', vertical='center')

    headers = ['Source DC'] + AGING_CATEGORIES + ['Total Pendency']
    data = []
    tot_buckets = defaultdict(int)

    for dc in ALLOWED_SOURCE_DCS:
        row_vals = [dc]
        row_tot = 0
        for cat in AGING_CATEGORIES:
            cnt = pivot[dc][cat]
            row_vals.append(cnt)
            row_tot += cnt
            tot_buckets[cat] += cnt
        row_vals.append(row_tot)
        data.append(row_vals)

    totals = ['Total'] + [tot_buckets[cat] for cat in AGING_CATEGORIES] + [sum(tot_buckets.values())]
    data.append(totals)

    end_col = len(headers)
    ws_sum.merge_cells(start_row=1, start_column=1, end_row=1, end_column=end_col)
    title_cell = ws_sum.cell(row=1, column=1, value='Aging wise report')
    title_cell.font = title_font
    title_cell.alignment = center
    for c in range(1, end_col + 1):
        cell = ws_sum.cell(row=1, column=c)
        cell.fill = title_fill
        cell.border = border

    for i, h in enumerate(headers):
        cell = ws_sum.cell(row=2, column=i + 1, value=h)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = center
        cell.border = border

    for r_idx, row_vals in enumerate(data):
        row_num_s = r_idx + 3
        is_total = row_vals[0] == 'Total'
        for c_idx, val in enumerate(row_vals):
            cell = ws_sum.cell(row=row_num_s, column=c_idx + 1, value=val)
            cell.border = border
            cell.alignment = left if c_idx == 0 else center
            if isinstance(val, (int, float)) and val > 0 and headers[c_idx] in ['3-5 Days', '6-10 Days', '>10 Days']:
                cell.fill = red_fill
                cell.font = red_font
            elif is_total:
                cell.fill = total_fill
                cell.font = total_font
            else:
                cell.font = data_font

    for col in range(1, end_col + 1):
        letter = get_column_letter(col)
        max_len = max(len(str(ws_sum.cell(r, col).value or '')) for r in range(1, len(data) + 3))
        ws_sum.column_dimensions[letter].width = max(max_len + 3, 14)

    temp_sum = io.BytesIO()
    out_wb.save(temp_sum)
    temp_sum.seek(0)
    out_wb.close()
    del out_wb, pivot
    gc.collect()

    # Assemble ZIP archive
    z_in = zipfile.ZipFile(temp_sum, 'r')
    z_out = zipfile.ZipFile(out_path, 'w', compression=zipfile.ZIP_DEFLATED)

    for item in z_in.infolist():
        if item.filename == '[Content_Types].xml':
            ct = z_in.read(item.filename).decode('utf-8')
            ct = ct.replace('</Types>', '<Override PartName="/xl/worksheets/sheet2.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/worksheets/sheet3.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/></Types>')
            z_out.writestr(item.filename, ct)
        elif item.filename == 'xl/workbook.xml':
            wb_xml = z_in.read(item.filename).decode('utf-8')
            wb_xml = wb_xml.replace('</sheets>', '<sheet name="Critical P0" sheetId="2" r:id="rId2"/><sheet name="Raw" sheetId="3" r:id="rId3"/></sheets>')
            z_out.writestr(item.filename, wb_xml)
        elif item.filename == 'xl/_rels/workbook.xml.rels':
            wb_rels = z_in.read(item.filename).decode('utf-8')
            wb_rels = wb_rels.replace('</Relationships>', '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet2.xml"/><Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet3.xml"/></Relationships>')
            z_out.writestr(item.filename, wb_rels)
        else:
            z_out.writestr(item, z_in.read(item.filename))

    with z_out.open('xl/worksheets/sheet2.xml', 'w', force_zip64=True) as zf_entry:
        with open(temp_p0_xml, 'rb') as f_p0:
            while True:
                buf = f_p0.read(1024 * 1024)
                if not buf:
                    break
                zf_entry.write(buf)

    with z_out.open('xl/worksheets/sheet3.xml', 'w', force_zip64=True) as zf_entry:
        with open(temp_raw_xml, 'rb') as f_raw:
            while True:
                buf = f_raw.read(1024 * 1024)
                if not buf:
                    break
                zf_entry.write(buf)

    z_out.close()
    z_in.close()

    if temp_raw_xml.exists():
        try:
            temp_raw_xml.unlink()
        except Exception:
            pass

    if temp_p0_xml.exists():
        try:
            temp_p0_xml.unlink()
        except Exception:
            pass

    gc.collect()
    log.info(f"Saved Reverse Pendency Report: {output_file.name} ({total_filtered} filtered rows, {p0_count} P0 rows)")
