"""
High-Speed Streaming Engine for ei_stream_server
================================================
Universal Zero-Memory Streaming Engine for reading and writing spreadsheets:
1. Reading: Streams rows one-by-one with constant O(1) memory (~25-35MB RAM).
2. Writing (XmlSheetWriter): Streams massive datasets directly to disk XML chunks with zero cell DOM overhead.
3. Assembly (assemble_stream_workbook): Stitches styled OpenPyXL Summary workbooks with disk-streamed XML data sheets.
4. Header Detection (ColumnFinder): Robust regex-based column detector resilient to spaces, casing, and underscores.
"""

import os
import io
import re
import csv
import math
import tempfile
import zipfile
import logging
from pathlib import Path
from typing import Iterator, List, Dict, Any, Optional, Union
import xml.etree.ElementTree as ET
from openpyxl.utils import get_column_letter

log = logging.getLogger("ei_stream_server.stream_engine")

try:
    from python_calamine import CalamineWorkbook
    HAS_CALAMINE = True
except ImportError:
    HAS_CALAMINE = False
    log.warning("python-calamine not installed; falling back to XML/CSV streaming.")


def esc(val: Any) -> str:
    """Fast XML entity escaping for Excel inline strings."""
    if val is None:
        return ""
    s = str(val)
    return s.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;").replace('"', "&quot;")


def _clean_key(name: Any) -> str:
    """Normalize header string by removing all whitespace, punctuation, and lowercasing."""
    if name is None:
        return ""
    return re.sub(r'[^a-z0-9]', '', str(name).lower())


class ColumnFinder:
    """
    Robust column detection helper supporting exact matches, cleaned keys, and substring fallbacks.
    """
    def __init__(self, headers: List[Any], schema: Optional[Dict[str, List[str]]] = None):
        self.headers = list(headers) if headers else []
        self.clean_map = {_clean_key(h): idx for idx, h in enumerate(self.headers) if h is not None}
        self.raw_map = {str(h).strip().lower(): idx for idx, h in enumerate(self.headers) if h is not None}
        self.resolved = {}

        if schema:
            for key, candidates in schema.items():
                self.resolved[key] = self.find(candidates)

    def find(self, candidates: List[str], fallback_keywords: Optional[List[str]] = None, default: int = 0) -> int:
        """Find index for given column candidates with fallback keywords."""
        if not self.headers:
            return default

        # 1. Exact cleaned match
        for cand in candidates:
            c_clean = _clean_key(cand)
            if c_clean in self.clean_map:
                return self.clean_map[c_clean]
            if cand.strip().lower() in self.raw_map:
                return self.raw_map[cand.strip().lower()]

        # 2. Substring / Keyword fallback
        keywords = fallback_keywords or candidates
        for kw in keywords:
            kw_clean = _clean_key(kw)
            for k, idx in self.clean_map.items():
                if kw_clean in k:
                    return idx

        return default

    def __getitem__(self, key: str) -> int:
        return self.resolved.get(key, 0)

    def get(self, key: str, default: int = 0) -> int:
        return self.resolved.get(key, default)


class XmlSheetWriter:
    """
    Context manager for streaming raw worksheet rows directly to a temporary XML file on disk.
    Holds virtually 0 MB in RAM regardless of row count.
    """
    def __init__(self, sheet_name: str, headers: Optional[List[Any]] = None, chunk_size: int = 1000):
        self.sheet_name = sheet_name
        self.headers = list(headers) if headers else []
        self.chunk_size = chunk_size
        self.row_counter = 1
        self.chunk: List[str] = []
        self.temp_file: Optional[Path] = None
        self._f: Optional[Any] = None

    def __enter__(self):
        temp_dir = Path(tempfile.gettempdir())
        safe_name = re.sub(r'[^a-zA-Z0-9_]', '_', self.sheet_name)
        self.temp_file = temp_dir / f"stream_{safe_name}_{os.getpid()}_{id(self)}.xml"
        self._f = open(self.temp_file, 'wb')
        self._f.write(b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData>')

        if self.headers:
            r_xml = ['<row r="1">']
            for c_i, h_val in enumerate(self.headers, 1):
                r_xml.append(f'<c r="{get_column_letter(c_i)}1" t="inlineStr"><is><t>{esc(h_val)}</t></is></c>')
            r_xml.append('</row>')
            self._f.write(''.join(r_xml).encode('utf-8'))
            self.row_counter = 2

        return self

    def write_row(self, row_values: Union[List[Any], tuple]):
        """Format and write a single row of values."""
        r_num = self.row_counter
        r_xml = [f'<row r="{r_num}">']
        
        for c_i, val in enumerate(row_values, 1):
            col_let = get_column_letter(c_i)
            if isinstance(val, (int, float)) and not math.isnan(val) and not math.isinf(val):
                r_xml.append(f'<c r="{col_let}{r_num}"><v>{val}</v></c>')
            else:
                r_xml.append(f'<c r="{col_let}{r_num}" t="inlineStr"><is><t>{esc(val)}</t></is></c>')

        r_xml.append('</row>')
        self.chunk.append(''.join(r_xml))
        self.row_counter += 1

        if len(self.chunk) >= self.chunk_size:
            self.flush()

    def flush(self):
        """Flush in-memory row chunk to disk."""
        if self.chunk and self._f:
            self._f.write(''.join(self.chunk).encode('utf-8'))
            self.chunk.clear()

    def close(self):
        """Finalize XML sheet structure and close disk file."""
        self.flush()
        if self._f and not self._f.closed:
            self._f.write(b'</sheetData></worksheet>')
            self._f.close()

    def cleanup(self):
        """Remove temporary XML disk file."""
        if self.temp_file and self.temp_file.exists():
            try:
                self.temp_file.unlink()
            except Exception:
                pass

    def __exit__(self, exc_type, exc_val, exc_tb):
        self.close()
        # Temp file is left for assemble_stream_workbook to read, will be cleaned in assemble_stream_workbook or cleanup


def assemble_stream_workbook(
    summary_workbook: Any,
    stream_sheets: List[XmlSheetWriter],
    output_path: Union[str, Path]
):
    """
    Stitches in-memory styled OpenPyXL Summary workbook with any number of disk-streamed XML data sheets.
    Generates standard valid OpenXML .xlsx archive with constant O(1) memory.
    """
    out_path = Path(output_path)
    out_path.parent.mkdir(parents=True, exist_ok=True)

    # Make sure all stream writers are closed
    for s_writer in stream_sheets:
        s_writer.close()

    # Save summary workbook to in-memory bytes
    temp_sum = io.BytesIO()
    summary_workbook.save(temp_sum)
    temp_sum.seek(0)
    try:
        summary_workbook.close()
    except Exception:
        pass

    # Read base zip archive and write final zip archive
    z_in = zipfile.ZipFile(temp_sum, 'r')
    z_out = zipfile.ZipFile(out_path, 'w', compression=zipfile.ZIP_DEFLATED)

    num_summary_sheets = len(summary_workbook.sheetnames) if hasattr(summary_workbook, 'sheetnames') else 1

    for item in z_in.infolist():
        if item.filename == '[Content_Types].xml':
            ct = z_in.read(item.filename).decode('utf-8')
            overrides = []
            for idx, _ in enumerate(stream_sheets, start=num_summary_sheets + 1):
                overrides.append(f'<Override PartName="/xl/worksheets/sheet{idx}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>')
            ct = ct.replace('</Types>', ''.join(overrides) + '</Types>')
            z_out.writestr(item.filename, ct)

        elif item.filename == 'xl/workbook.xml':
            wb_xml = z_in.read(item.filename).decode('utf-8')
            sheet_tags = []
            for idx, s_writer in enumerate(stream_sheets, start=num_summary_sheets + 1):
                safe_sheet_name = esc(s_writer.sheet_name)
                sheet_tags.append(f'<sheet name="{safe_sheet_name}" sheetId="{idx}" r:id="rId{idx}"/>')
            wb_xml = wb_xml.replace('</sheets>', ''.join(sheet_tags) + '</sheets>')
            z_out.writestr(item.filename, wb_xml)

        elif item.filename == 'xl/_rels/workbook.xml.rels':
            wb_rels = z_in.read(item.filename).decode('utf-8')
            rel_tags = []
            for idx, _ in enumerate(stream_sheets, start=num_summary_sheets + 1):
                rel_tags.append(f'<Relationship Id="rId{idx}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet{idx}.xml"/>')
            wb_rels = wb_rels.replace('</Relationships>', ''.join(rel_tags) + '</Relationships>')
            z_out.writestr(item.filename, wb_rels)

        else:
            z_out.writestr(item, z_in.read(item.filename))

    # Stream each XML data sheet from disk directly into the ZIP entry
    for idx, s_writer in enumerate(stream_sheets, start=num_summary_sheets + 1):
        with z_out.open(f'xl/worksheets/sheet{idx}.xml', 'w', force_zip64=True) as zf_entry:
            if s_writer.temp_file and s_writer.temp_file.exists():
                with open(s_writer.temp_file, 'rb') as f_xml_in:
                    while True:
                        buf = f_xml_in.read(1024 * 1024)
                        if not buf:
                            break
                        zf_entry.write(buf)

    z_out.close()
    z_in.close()

    # Clean up all temp XML files
    for s_writer in stream_sheets:
        s_writer.cleanup()

    log.info(f"Successfully assembled streaming workbook: {out_path.name} with {len(stream_sheets)} streamed sheet(s)")


# =========================================================================
# READING FUNCTIONS
# =========================================================================

def get_sheet_names(file_path: Union[str, Path]) -> List[str]:
    """Return all worksheet names in a spreadsheet file using 0 memory."""
    path = Path(file_path)
    ext = path.suffix.lower()

    if HAS_CALAMINE and ext in ('.xlsx', '.xlsb', '.ods', '.xls', '.xlsm'):
        try:
            wb = CalamineWorkbook.from_path(str(path))
            return wb.sheet_names
        except Exception as e:
            log.warning(f"Calamine get_sheet_names failed ({e}); falling back...")

    if ext in ('.xlsx', '.xlsm') and zipfile.is_zipfile(path):
        try:
            with zipfile.ZipFile(path, 'r') as z:
                if 'xl/workbook.xml' in z.namelist():
                    with z.open('xl/workbook.xml') as f:
                        tree = ET.parse(f)
                        root = tree.getroot()
                        ns = {'ns': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}
                        sheets = root.findall('.//ns:sheet', ns) or root.findall('.//{http://schemas.openxmlformats.org/spreadsheetml/2006/main}sheet')
                        if sheets:
                            return [s.attrib.get('name', '') for s in sheets if s.attrib.get('name')]
        except Exception as e:
            log.warning(f"Zip workbook parse failed: {e}")

    if ext in ('.csv', '.tsv', '.txt'):
        return [path.stem]

    try:
        import openpyxl
        wb = openpyxl.load_workbook(str(path), read_only=True)
        names = list(wb.sheetnames)
        wb.close()
        return names
    except Exception as e:
        log.error(f"Failed to extract sheet names from {path}: {e}")
        return ["Sheet1"]


def stream_sheet_rows(
    file_path: Union[str, Path],
    sheet_name: Optional[str] = None,
    start_row: int = 1
) -> Iterator[List[Any]]:
    """
    Stream rows one-by-one from a spreadsheet.
    Yields list of cell values for each row without loading the entire sheet into memory.
    """
    path = Path(file_path)
    ext = path.suffix.lower()

    # 1. Primary: High-speed Rust Calamine (Supports .xlsx, .xlsb, .ods, .xls, .xlsm)
    if HAS_CALAMINE and ext in ('.xlsx', '.xlsb', '.ods', '.xls', '.xlsm'):
        try:
            wb = CalamineWorkbook.from_path(str(path))
            target_sheet = sheet_name or (wb.sheet_names[0] if wb.sheet_names else "Sheet1")
            
            if target_sheet not in wb.sheet_names:
                target_clean = str(target_sheet).lower().replace(' ', '_').replace('-', '_')
                found = False
                for s in wb.sheet_names:
                    s_clean = s.lower().replace(' ', '_').replace('-', '_')
                    if s_clean == target_clean or target_clean in s_clean:
                        target_sheet = s
                        found = True
                        break
                if not found and wb.sheet_names:
                    target_sheet = wb.sheet_names[0]

            sheet = wb.get_sheet_by_name(target_sheet)
            current_row = 1
            for row in sheet.iter_rows():
                if current_row >= start_row:
                    yield list(row)
                current_row += 1
            return
        except Exception as e:
            log.warning(f"Calamine stream failed for {path.name} ({e}). Falling back to openpyxl...")

    # 2. Fallback for CSV / TSV / Text files
    if ext in ('.csv', '.tsv', '.txt'):
        for sep in [',', '\t', ';', '|']:
            for enc in ['utf-8-sig', 'utf-8', 'latin1', 'cp1252']:
                try:
                    with open(path, 'r', encoding=enc, errors='replace') as f:
                        reader = csv.reader(f, delimiter=sep)
                        current_row = 1
                        for row in reader:
                            if current_row >= start_row:
                                yield row
                            current_row += 1
                        return
                except Exception:
                    continue

    # 3. Fallback: openpyxl read_only stream
    try:
        import openpyxl
        wb = openpyxl.load_workbook(str(path), read_only=True, data_only=True)
        target_sheet = sheet_name or wb.sheetnames[0]
        if target_sheet not in wb.sheetnames:
            target_clean = str(target_sheet).lower().replace(' ', '_').replace('-', '_')
            found = False
            for s in wb.sheetnames:
                s_clean = s.lower().replace(' ', '_').replace('-', '_')
                if s_clean == target_clean or target_clean in s_clean:
                    target_sheet = s
                    found = True
                    break
            if not found and wb.sheetnames:
                target_sheet = wb.sheetnames[0]

        ws = wb[target_sheet]
        current_row = 1
        for row in ws.iter_rows(values_only=True):
            if current_row >= start_row:
                yield list(row) if row else []
            current_row += 1
        wb.close()
    except Exception as e:
        log.error(f"Failed to stream rows from {path.name}: {e}")
        raise


def stream_sheet_dicts(
    file_path: Union[str, Path],
    sheet_name: Optional[str] = None,
    header_row: int = 1
) -> Iterator[Dict[str, Any]]:
    """Yields dictionary for each row mapped to the header names defined in header_row."""
    row_iter = stream_sheet_rows(file_path, sheet_name=sheet_name, start_row=header_row)
    try:
        raw_headers = next(row_iter)
    except StopIteration:
        return

    headers = [str(h).strip() if h is not None else f"col_{i}" for i, h in enumerate(raw_headers)]

    for row in row_iter:
        if not row or all(v is None or str(v).strip() == "" for v in row):
            continue
        row_dict = {}
        for i, val in enumerate(row):
            col_name = headers[i] if i < len(headers) else f"col_{i}"
            row_dict[col_name] = val
        yield row_dict


def inspect_spreadsheet_headers(file_path: Union[str, Path], sheet_name: Optional[str] = None) -> List[str]:
    """Inspect first non-empty row of a sheet to detect column headers without loading the file."""
    rows_iter = stream_sheet_rows(file_path, sheet_name=sheet_name, start_row=1)
    for row in rows_iter:
        if row and any(cell is not None and str(cell).strip() != "" for cell in row):
            return [str(c).strip() if c is not None else "" for c in row]
    return []
