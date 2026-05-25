#!/usr/bin/env python3
"""Generate Tabel_Premi_MiUHC.xlsx from PDF source files."""
import re
import zlib
import zipfile
import xml.etree.ElementTree as ET
from io import BytesIO
import os

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
KONV_PDF = os.path.join(SCRIPT_DIR, 'B. Tabel Premi Miuhc Versi Quick PDF Per Sep 2025 (3).pdf')
SYARIAH_PDF = os.path.join(SCRIPT_DIR, 'B. Tabel Premi Syariah 2024 - Miuhc Versi Quick PDF (3).pdf')
OUTPUT_XLSX = os.path.join(SCRIPT_DIR, 'Tabel_Premi_MiUHC.xlsx')

PLANS_REGULAR = ["Diamond", "Ruby", "Emerald", "Topaz", "Topaz ID", "Jade", "Jade ID", "Sapphire"]
PLANS_SMART_KONV = ["Diamond Smart", "Ruby Smart", "Emerald Smart", "Topaz Smart", "Topaz ID Smart", "Jade Smart", "Jade ID Smart"]
PLANS_SMART_SYARIAH = ["Diamond Smart", "Ruby Smart", "Emerald Smart", "Topaz Smart", "Jade Smart"]

WA_LINK = "https://wa.me/6287781896087"
HEADER_TEXT = "PETRUS PARTNER MANULIFE 087781896087"


def extract_text_streams(pdf_path):
    """Extract text from all zlib-compressed streams in a PDF."""
    with open(pdf_path, "rb") as f:
        data = f.read()
    all_text_streams = []
    for match in re.finditer(rb'stream\r?\n(.*?)\r?\nendstream', data, re.DOTALL):
        raw = match.group(1)
        try:
            decompressed = zlib.decompress(raw)
        except Exception:
            continue
        text_parts = []
        for m in re.finditer(rb'\[(.*?)\]\s*TJ', decompressed, re.DOTALL):
            parts = re.findall(rb'\((.*?)\)', m.group(1))
            text_parts.append(b"".join(parts).decode("latin-1", errors="replace"))
        for m in re.finditer(rb'\((.*?)\)\s*Tj', decompressed):
            text_parts.append(m.group(1).decode("latin-1", errors="replace"))
        if text_parts:
            all_text_streams.append(text_parts)
    return all_text_streams


def extract_numbers(text_parts):
    """Extract comma-formatted numbers from text parts list."""
    nums = []
    for p in text_parts:
        if re.match(r'\d{1,3}(,\d{3})+$', p):
            nums.append(p)
    return nums


def parse_konvensional():
    """Parse Konvensional PDF data. Returns list of dicts per age."""
    text_streams = extract_text_streams(KONV_PDF)
    ages_data = []
    # Text streams 2..87 = ages 0..85
    for age in range(86):
        stream_idx = age + 2
        if stream_idx >= len(text_streams):
            break
        parts = text_streams[stream_idx]
        nums = extract_numbers(parts)
        full_text = ' '.join(parts)
        has_rawat_jalan = 'Rawat Jalan' in full_text

        entry = {'age': age, 'pria': {}, 'wanita': {}}
        if has_rawat_jalan:
            # 60 numbers: PRIA (8+7+8+7=30) then WANITA (8+7+8+7=30)
            entry['pria']['hospital'] = nums[0:8]
            entry['pria']['hospital_smart'] = nums[8:15]
            entry['pria']['outpatient'] = nums[15:23]
            entry['pria']['outpatient_smart'] = nums[23:30]
            entry['wanita']['hospital'] = nums[30:38]
            entry['wanita']['hospital_smart'] = nums[38:45]
            entry['wanita']['outpatient'] = nums[45:53]
            entry['wanita']['outpatient_smart'] = nums[53:60]
        else:
            # 30 numbers: PRIA (8+7=15) then WANITA (8+7=15) - Hospital only
            entry['pria']['hospital'] = nums[0:8]
            entry['pria']['hospital_smart'] = nums[8:15]
            entry['pria']['outpatient'] = []
            entry['pria']['outpatient_smart'] = []
            entry['wanita']['hospital'] = nums[15:23]
            entry['wanita']['hospital_smart'] = nums[23:30]
            entry['wanita']['outpatient'] = []
            entry['wanita']['outpatient_smart'] = []
        ages_data.append(entry)
    return ages_data


def parse_syariah():
    """Parse Syariah PDF data. Returns list of dicts per age."""
    text_streams = extract_text_streams(SYARIAH_PDF)
    ages_data = []
    # Text streams 3..174 = alternating PRIA/WANITA for ages 0..85
    for age in range(86):
        pria_idx = 3 + age * 2
        wanita_idx = 4 + age * 2
        if pria_idx >= len(text_streams) or wanita_idx >= len(text_streams):
            break

        pria_parts = text_streams[pria_idx]
        wanita_parts = text_streams[wanita_idx]
        pria_nums = extract_numbers(pria_parts)
        wanita_nums = extract_numbers(wanita_parts)

        pria_full = ' '.join(pria_parts)
        has_rawat = 'Rawat Jalan' in pria_full
        has_gigi = 'Gigi' in pria_full

        entry = {'age': age, 'pria': {}, 'wanita': {}}

        if has_gigi:
            # 39 numbers: Hospital(8) + Hospital Smart(5) + Outpatient(8) + Outpatient Smart(5) + Dental(8) + Dental Smart(5)
            entry['pria']['hospital'] = pria_nums[0:8]
            entry['pria']['hospital_smart'] = pria_nums[8:13]
            entry['pria']['outpatient'] = pria_nums[13:21]
            entry['pria']['outpatient_smart'] = pria_nums[21:26]
            entry['pria']['dental'] = pria_nums[26:34]
            entry['pria']['dental_smart'] = pria_nums[34:39]
            entry['wanita']['hospital'] = wanita_nums[0:8]
            entry['wanita']['hospital_smart'] = wanita_nums[8:13]
            entry['wanita']['outpatient'] = wanita_nums[13:21]
            entry['wanita']['outpatient_smart'] = wanita_nums[21:26]
            entry['wanita']['dental'] = wanita_nums[26:34]
            entry['wanita']['dental_smart'] = wanita_nums[34:39]
        elif has_rawat:
            # Has Hospital + Outpatient but no Dental
            entry['pria']['hospital'] = pria_nums[0:8]
            entry['pria']['hospital_smart'] = pria_nums[8:13]
            entry['pria']['outpatient'] = pria_nums[13:21]
            entry['pria']['outpatient_smart'] = pria_nums[21:26]
            entry['pria']['dental'] = []
            entry['pria']['dental_smart'] = []
            entry['wanita']['hospital'] = wanita_nums[0:8]
            entry['wanita']['hospital_smart'] = wanita_nums[8:13]
            entry['wanita']['outpatient'] = wanita_nums[13:21]
            entry['wanita']['outpatient_smart'] = wanita_nums[21:26]
            entry['wanita']['dental'] = []
            entry['wanita']['dental_smart'] = []
        else:
            # Hospital only (ages 80-85)
            entry['pria']['hospital'] = pria_nums[0:8]
            entry['pria']['hospital_smart'] = pria_nums[8:13]
            entry['pria']['outpatient'] = []
            entry['pria']['outpatient_smart'] = []
            entry['pria']['dental'] = []
            entry['pria']['dental_smart'] = []
            entry['wanita']['hospital'] = wanita_nums[0:8]
            entry['wanita']['hospital_smart'] = wanita_nums[8:13]
            entry['wanita']['outpatient'] = []
            entry['wanita']['outpatient_smart'] = []
            entry['wanita']['dental'] = []
            entry['wanita']['dental_smart'] = []
        ages_data.append(entry)
    return ages_data


class XlsxWriter:
    """Build an xlsx file from scratch using zipfile and XML."""

    def __init__(self):
        self.sheets = []  # list of (name, rows)
        self.shared_strings = []
        self.ss_map = {}  # string -> index
        self.hyperlinks = {}  # sheet_idx -> [(cell_ref, url)]

    def add_sheet(self, name, rows):
        """Add a sheet. rows is list of row-lists. Each cell is (value, style_id)."""
        self.sheets.append((name, rows))

    def _get_ss_index(self, s):
        """Get shared string index, adding if needed."""
        if s not in self.ss_map:
            self.ss_map[s] = len(self.shared_strings)
            self.shared_strings.append(s)
        return self.ss_map[s]

    def _col_letter(self, col_idx):
        """Convert 0-based column index to Excel column letter(s)."""
        result = ""
        idx = col_idx
        while True:
            result = chr(65 + idx % 26) + result
            idx = idx // 26 - 1
            if idx < 0:
                break
        return result

    def _cell_ref(self, row, col):
        """Get cell reference like A1, B2, etc."""
        return f"{self._col_letter(col)}{row + 1}"

    def add_hyperlink(self, sheet_idx, row, col, url):
        """Add a hyperlink to a cell."""
        if sheet_idx not in self.hyperlinks:
            self.hyperlinks[sheet_idx] = []
        self.hyperlinks[sheet_idx].append((self._cell_ref(row, col), url))

    def _build_content_types(self):
        ct = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        ct += '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
        ct += '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
        ct += '<Default Extension="xml" ContentType="application/xml"/>'
        ct += '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'
        ct += '<Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>'
        ct += '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>'
        for i in range(len(self.sheets)):
            ct += f'<Override PartName="/xl/worksheets/sheet{i+1}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
        ct += '</Types>'
        return ct

    def _build_rels(self):
        r = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        r += '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        r += '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>'
        r += '</Relationships>'
        return r

    def _build_workbook(self):
        wb = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        wb += '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
        wb += '<sheets>'
        for i, (name, _) in enumerate(self.sheets):
            wb += f'<sheet name="{name}" sheetId="{i+1}" r:id="rId{i+1}"/>'
        wb += '</sheets></workbook>'
        return wb

    def _build_workbook_rels(self):
        r = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        r += '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        for i in range(len(self.sheets)):
            r += f'<Relationship Id="rId{i+1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet{i+1}.xml"/>'
        r += f'<Relationship Id="rId{len(self.sheets)+1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/>'
        r += f'<Relationship Id="rId{len(self.sheets)+2}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>'
        r += '</Relationships>'
        return r

    def _build_styles(self):
        # Styles:
        # 0 = default
        # 1 = bold (header)
        # 2 = green bg white text bold (section header)
        # 3 = bold green text (branding header)
        # 4 = right-aligned (numbers)
        s = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        s += '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
        # Fonts
        s += '<fonts count="4">'
        s += '<font><sz val="11"/><name val="Calibri"/></font>'  # 0: default
        s += '<font><b/><sz val="11"/><name val="Calibri"/></font>'  # 1: bold
        s += '<font><b/><sz val="11"/><color rgb="FFFFFFFF"/><name val="Calibri"/></font>'  # 2: bold white
        s += '<font><b/><sz val="11"/><color rgb="FF00B050"/><name val="Calibri"/></font>'  # 3: bold green
        s += '</fonts>'
        # Fills
        s += '<fills count="3">'
        s += '<fill><patternFill patternType="none"/></fill>'  # 0
        s += '<fill><patternFill patternType="gray125"/></fill>'  # 1
        s += '<fill><patternFill patternType="solid"><fgColor rgb="FF00B050"/></patternFill></fill>'  # 2: green
        s += '</fills>'
        # Borders
        s += '<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>'
        # CellStyleXfs
        s += '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>'
        # CellXfs
        s += '<cellXfs count="5">'
        s += '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>'  # 0: default
        s += '<xf numFmtId="0" fontId="1" fillId="0" borderId="0" xfId="0" applyFont="1"/>'  # 1: bold
        s += '<xf numFmtId="0" fontId="2" fillId="2" borderId="0" xfId="0" applyFont="1" applyFill="1"/>'  # 2: green bg white bold
        s += '<xf numFmtId="0" fontId="3" fillId="0" borderId="0" xfId="0" applyFont="1"/>'  # 3: bold green text
        s += '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyAlignment="1"><alignment horizontal="right"/></xf>'  # 4: right-aligned
        s += '</cellXfs>'
        s += '</styleSheet>'
        return s

    def _build_shared_strings(self):
        ss = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        ss += f'<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="{len(self.shared_strings)}" uniqueCount="{len(self.shared_strings)}">'
        for s in self.shared_strings:
            escaped = s.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
            ss += f'<si><t>{escaped}</t></si>'
        ss += '</sst>'
        return ss

    def _build_sheet(self, sheet_idx):
        _, rows = self.sheets[sheet_idx]
        s = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        s += '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
        # Column widths
        s += '<cols>'
        s += '<col min="1" max="1" width="32" customWidth="1"/>'
        s += '<col min="2" max="9" width="16" customWidth="1"/>'
        s += '</cols>'
        s += '<sheetData>'
        for r_idx, row in enumerate(rows):
            s += f'<row r="{r_idx + 1}">'
            for c_idx, cell in enumerate(row):
                if cell is None:
                    continue
                value, style_id = cell
                ref = self._cell_ref(r_idx, c_idx)
                si = self._get_ss_index(str(value))
                s += f'<c r="{ref}" t="s" s="{style_id}"><v>{si}</v></c>'
            s += '</row>'
        s += '</sheetData>'
        # Hyperlinks
        if sheet_idx in self.hyperlinks:
            s += '<hyperlinks>'
            for i, (cell_ref, _) in enumerate(self.hyperlinks[sheet_idx]):
                s += f'<hyperlink ref="{cell_ref}" r:id="rId{i+1}"/>'
            s += '</hyperlinks>'
        s += '</worksheet>'
        return s

    def _build_sheet_rels(self, sheet_idx):
        if sheet_idx not in self.hyperlinks or not self.hyperlinks[sheet_idx]:
            return None
        r = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        r += '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        for i, (_, url) in enumerate(self.hyperlinks[sheet_idx]):
            r += f'<Relationship Id="rId{i+1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="{url}" TargetMode="External"/>'
        r += '</Relationships>'
        return r

    def save(self, filepath):
        """Save the xlsx file."""
        # First pass: build sheets to populate shared strings
        sheet_xmls = []
        for i in range(len(self.sheets)):
            sheet_xmls.append(self._build_sheet(i))

        with zipfile.ZipFile(filepath, 'w', zipfile.ZIP_DEFLATED) as zf:
            zf.writestr('[Content_Types].xml', self._build_content_types())
            zf.writestr('_rels/.rels', self._build_rels())
            zf.writestr('xl/workbook.xml', self._build_workbook())
            zf.writestr('xl/_rels/workbook.xml.rels', self._build_workbook_rels())
            zf.writestr('xl/styles.xml', self._build_styles())
            zf.writestr('xl/sharedStrings.xml', self._build_shared_strings())
            for i in range(len(self.sheets)):
                zf.writestr(f'xl/worksheets/sheet{i+1}.xml', sheet_xmls[i])
                rels = self._build_sheet_rels(i)
                if rels:
                    zf.writestr(f'xl/worksheets/_rels/sheet{i+1}.xml.rels', rels)


def build_konv_sheet(ages_data):
    """Build rows for Konvensional sheet."""
    rows = []
    # Row 0: Header with WhatsApp link
    row = [(HEADER_TEXT, 3)]
    rows.append(row)
    # Row 1: Title
    row = [("Premi MiUltimate HealthCare 2025", 1)]
    rows.append(row)
    # Row 2: Empty
    rows.append([])

    for entry in ages_data:
        age = entry['age']
        # PRIA section
        rows.append([(f"PRIA, {age} TAHUN", 1)])
        # Hospital
        rows.append([("Manfaat Perawatan Rumah Sakit", 2)])
        # Plan names
        row = [(None, 0)] + [(p, 1) for p in PLANS_REGULAR]
        rows.append(row)
        # Values
        row = [(None, 0)] + [(v, 4) for v in entry['pria']['hospital']]
        rows.append(row)
        # Smart plan names
        row = [(None, 0)] + [(p, 1) for p in PLANS_SMART_KONV]
        rows.append(row)
        # Smart values
        row = [(None, 0)] + [(v, 4) for v in entry['pria']['hospital_smart']]
        rows.append(row)

        # Outpatient (if available)
        if entry['pria']['outpatient']:
            rows.append([("Manfaat Rawat Jalan", 2)])
            row = [(None, 0)] + [(p, 1) for p in PLANS_REGULAR]
            rows.append(row)
            row = [(None, 0)] + [(v, 4) for v in entry['pria']['outpatient']]
            rows.append(row)
            row = [(None, 0)] + [(p, 1) for p in PLANS_SMART_KONV]
            rows.append(row)
            row = [(None, 0)] + [(v, 4) for v in entry['pria']['outpatient_smart']]
            rows.append(row)

        # WANITA section
        rows.append([(f"WANITA, {age} TAHUN", 1)])
        # Hospital
        rows.append([("Manfaat Perawatan Rumah Sakit", 2)])
        row = [(None, 0)] + [(p, 1) for p in PLANS_REGULAR]
        rows.append(row)
        row = [(None, 0)] + [(v, 4) for v in entry['wanita']['hospital']]
        rows.append(row)
        row = [(None, 0)] + [(p, 1) for p in PLANS_SMART_KONV]
        rows.append(row)
        row = [(None, 0)] + [(v, 4) for v in entry['wanita']['hospital_smart']]
        rows.append(row)

        # Outpatient (if available)
        if entry['wanita']['outpatient']:
            rows.append([("Manfaat Rawat Jalan", 2)])
            row = [(None, 0)] + [(p, 1) for p in PLANS_REGULAR]
            rows.append(row)
            row = [(None, 0)] + [(v, 4) for v in entry['wanita']['outpatient']]
            rows.append(row)
            row = [(None, 0)] + [(p, 1) for p in PLANS_SMART_KONV]
            rows.append(row)
            row = [(None, 0)] + [(v, 4) for v in entry['wanita']['outpatient_smart']]
            rows.append(row)

        # Separator
        rows.append([])

    return rows


def build_syariah_sheet(ages_data):
    """Build rows for Syariah sheet."""
    rows = []
    # Row 0: Header with WhatsApp link
    row = [(HEADER_TEXT, 3)]
    rows.append(row)
    # Row 1: Title
    row = [("Premi MiUltimate HealthCare 2024", 1)]
    rows.append(row)
    # Row 2: Empty
    rows.append([])

    for entry in ages_data:
        age = entry['age']
        # PRIA section
        rows.append([(f"PRIA, {age} TAHUN", 1)])
        # Hospital
        rows.append([("Manfaat Perawatan Rumah Sakit", 2)])
        row = [(None, 0)] + [(p, 1) for p in PLANS_REGULAR]
        rows.append(row)
        row = [(None, 0)] + [(v, 4) for v in entry['pria']['hospital']]
        rows.append(row)
        row = [(None, 0)] + [(p, 1) for p in PLANS_SMART_SYARIAH]
        rows.append(row)
        row = [(None, 0)] + [(v, 4) for v in entry['pria']['hospital_smart']]
        rows.append(row)

        # Outpatient (if available)
        if entry['pria']['outpatient']:
            rows.append([("Manfaat Rawat Jalan", 2)])
            row = [(None, 0)] + [(p, 1) for p in PLANS_REGULAR]
            rows.append(row)
            row = [(None, 0)] + [(v, 4) for v in entry['pria']['outpatient']]
            rows.append(row)
            row = [(None, 0)] + [(p, 1) for p in PLANS_SMART_SYARIAH]
            rows.append(row)
            row = [(None, 0)] + [(v, 4) for v in entry['pria']['outpatient_smart']]
            rows.append(row)

        # Dental (if available)
        if entry['pria']['dental']:
            rows.append([("Manfaat Perawatan Gigi", 2)])
            row = [(None, 0)] + [(p, 1) for p in PLANS_REGULAR]
            rows.append(row)
            row = [(None, 0)] + [(v, 4) for v in entry['pria']['dental']]
            rows.append(row)
            row = [(None, 0)] + [(p, 1) for p in PLANS_SMART_SYARIAH]
            rows.append(row)
            row = [(None, 0)] + [(v, 4) for v in entry['pria']['dental_smart']]
            rows.append(row)

        # WANITA section
        rows.append([(f"WANITA, {age} TAHUN", 1)])
        # Hospital
        rows.append([("Manfaat Perawatan Rumah Sakit", 2)])
        row = [(None, 0)] + [(p, 1) for p in PLANS_REGULAR]
        rows.append(row)
        row = [(None, 0)] + [(v, 4) for v in entry['wanita']['hospital']]
        rows.append(row)
        row = [(None, 0)] + [(p, 1) for p in PLANS_SMART_SYARIAH]
        rows.append(row)
        row = [(None, 0)] + [(v, 4) for v in entry['wanita']['hospital_smart']]
        rows.append(row)

        # Outpatient (if available)
        if entry['wanita']['outpatient']:
            rows.append([("Manfaat Rawat Jalan", 2)])
            row = [(None, 0)] + [(p, 1) for p in PLANS_REGULAR]
            rows.append(row)
            row = [(None, 0)] + [(v, 4) for v in entry['wanita']['outpatient']]
            rows.append(row)
            row = [(None, 0)] + [(p, 1) for p in PLANS_SMART_SYARIAH]
            rows.append(row)
            row = [(None, 0)] + [(v, 4) for v in entry['wanita']['outpatient_smart']]
            rows.append(row)

        # Dental (if available)
        if entry['wanita']['dental']:
            rows.append([("Manfaat Perawatan Gigi", 2)])
            row = [(None, 0)] + [(p, 1) for p in PLANS_REGULAR]
            rows.append(row)
            row = [(None, 0)] + [(v, 4) for v in entry['wanita']['dental']]
            rows.append(row)
            row = [(None, 0)] + [(p, 1) for p in PLANS_SMART_SYARIAH]
            rows.append(row)
            row = [(None, 0)] + [(v, 4) for v in entry['wanita']['dental_smart']]
            rows.append(row)

        # Separator
        rows.append([])

    return rows


def main():
    print("Parsing Konvensional PDF...")
    konv_data = parse_konvensional()
    print(f"  Extracted data for {len(konv_data)} ages")

    print("Parsing Syariah PDF...")
    syariah_data = parse_syariah()
    print(f"  Extracted data for {len(syariah_data)} ages")

    print("Building Excel file...")
    writer = XlsxWriter()

    konv_rows = build_konv_sheet(konv_data)
    syariah_rows = build_syariah_sheet(syariah_data)

    writer.add_sheet("Konvensional", konv_rows)
    writer.add_sheet("Syariah", syariah_rows)

    # Add hyperlinks for row 0 col 0 on both sheets
    writer.add_hyperlink(0, 0, 0, WA_LINK)
    writer.add_hyperlink(1, 0, 0, WA_LINK)

    writer.save(OUTPUT_XLSX)
    print(f"Excel file saved to: {OUTPUT_XLSX}")

    # Verify
    import zipfile as zf_verify
    with zf_verify.ZipFile(OUTPUT_XLSX) as z:
        print(f"  Valid xlsx with {len(z.namelist())} files")
        print(f"  Contents: {z.namelist()}")


if __name__ == "__main__":
    main()
