#!/usr/bin/env python3
"""Generate Tabel_Premi_MiUHC.xlsx from PDF source files."""
import re
import zlib
import zipfile
import os

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
KONV_PDF = os.path.join(SCRIPT_DIR, 'B. Tabel Premi Miuhc Versi Quick PDF Per Sep 2025 (3).pdf')
SYARIAH_PDF = os.path.join(SCRIPT_DIR, 'B. Tabel Premi Syariah 2024 - Miuhc Versi Quick PDF (3).pdf')
OUTPUT_XLSX = os.path.join(SCRIPT_DIR, 'Tabel_Premi_MiUHC.xlsx')

PLANS_REGULAR = ["Diamond", "Ruby", "Emerald", "Topaz", "Topaz ID", "Jade", "Jade ID", "Sapphire"]
PLANS_SMART_KONV = ["Diamond Smart", "Ruby Smart", "Emerald Smart", "Topaz Smart", "Topaz ID Smart", "Jade Smart", "Jade ID Smart"]
PLANS_SMART_SYARIAH = ["Diamond Smart", "Ruby Smart", "Emerald Smart", "Topaz Smart", "Jade Smart"]

WA_LINK = "https://wa.me/6287781896087"
HEADER_TEXT = "PETRUS JAKUB MANULIFE 087781896087"


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
            entry['pria']['hospital'] = nums[0:8]
            entry['pria']['hospital_smart'] = nums[8:15]
            entry['pria']['outpatient'] = nums[15:23]
            entry['pria']['outpatient_smart'] = nums[23:30]
            entry['wanita']['hospital'] = nums[30:38]
            entry['wanita']['hospital_smart'] = nums[38:45]
            entry['wanita']['outpatient'] = nums[45:53]
            entry['wanita']['outpatient_smart'] = nums[53:60]
        else:
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


def _num(s):
    """Convert comma-formatted number string to integer."""
    return int(s.replace(",", ""))


def _col_letter(col_idx):
    """Convert 0-based column index to Excel column letter(s)."""
    result = ""
    idx = col_idx
    while True:
        result = chr(65 + idx % 26) + result
        idx = idx // 26 - 1
        if idx < 0:
            break
    return result


def _cell_ref(row, col):
    """Get cell reference like A1, B2, etc. (0-based row and col)."""
    return f"{_col_letter(col)}{row + 1}"


def _xml_escape(s):
    """Escape text for XML content."""
    return s.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;").replace('"', "&quot;")


class XlsxWriter:
    """Build an xlsx file from scratch using zipfile and XML."""

    def __init__(self):
        self.sheets = []  # list of (name, rows, data_validations, merge_cells)
        self.shared_strings = []
        self.ss_map = {}  # string -> index
        self.hyperlinks = {}  # sheet_idx -> [(cell_ref, url)]

    def add_sheet(self, name, rows, data_validations=None, merge_cells=None):
        """Add a sheet.
        rows: list of row-lists. Each cell is:
          (value, style_id) or (value, style_id, 'string') - shared string
          (value, style_id, 'number') - numeric cell
          (value, style_id, 'formula') - formula cell
          None - empty cell
        data_validations: list of (sqref, formula1_str) e.g. [("B4", '"Pria,Wanita"')]
        merge_cells: list of ref strings e.g. ["A1:I1"]
        """
        self.sheets.append((name, rows, data_validations or [], merge_cells or []))

    def _get_ss_index(self, s):
        """Get shared string index, adding if needed."""
        if s not in self.ss_map:
            self.ss_map[s] = len(self.shared_strings)
            self.shared_strings.append(s)
        return self.ss_map[s]

    def add_hyperlink(self, sheet_idx, row, col, url):
        """Add a hyperlink to a cell."""
        if sheet_idx not in self.hyperlinks:
            self.hyperlinks[sheet_idx] = []
        self.hyperlinks[sheet_idx].append((_cell_ref(row, col), url))

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
        for i, (name, _, _, _) in enumerate(self.sheets):
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
        # 1 = bold
        # 2 = green bg white text bold
        # 3 = bold green text
        # 4 = right-aligned
        # 5 = number format #,##0 (for formula results and data table numbers)
        s = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        s += '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
        # Number formats
        s += '<numFmts count="1">'
        s += '<numFmt numFmtId="164" formatCode="#,##0"/>'
        s += '</numFmts>'
        # Fonts
        s += '<fonts count="4">'
        s += '<font><sz val="11"/><name val="Calibri"/></font>'  # 0: default
        s += '<font><b/><sz val="11"/><name val="Calibri"/></font>'  # 1: bold
        s += '<font><b/><sz val="11"/><color rgb="FFFFFFFF"/><name val="Calibri"/></font>'  # 2: bold white
        s += '<font><b/><sz val="11"/><color rgb="FF00B050"/><name val="Calibri"/></font>'  # 3: bold green
        s += '</fonts>'
        # Fills
        s += '<fills count="3">'
        s += '<fill><patternFill patternType="none"/></fill>'
        s += '<fill><patternFill patternType="gray125"/></fill>'
        s += '<fill><patternFill patternType="solid"><fgColor rgb="FF00B050"/></patternFill></fill>'  # 2: green
        s += '</fills>'
        # Borders
        s += '<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>'
        # CellStyleXfs
        s += '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>'
        # CellXfs
        s += '<cellXfs count="6">'
        s += '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>'  # 0: default
        s += '<xf numFmtId="0" fontId="1" fillId="0" borderId="0" xfId="0" applyFont="1"/>'  # 1: bold
        s += '<xf numFmtId="0" fontId="2" fillId="2" borderId="0" xfId="0" applyFont="1" applyFill="1"/>'  # 2: green bg white bold
        s += '<xf numFmtId="0" fontId="3" fillId="0" borderId="0" xfId="0" applyFont="1"/>'  # 3: bold green text
        s += '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyAlignment="1"><alignment horizontal="right"/></xf>'  # 4: right-aligned
        s += '<xf numFmtId="164" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>'  # 5: number format #,##0
        s += '</cellXfs>'
        s += '</styleSheet>'
        return s

    def _build_shared_strings(self):
        ss = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        ss += f'<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="{len(self.shared_strings)}" uniqueCount="{len(self.shared_strings)}">'
        for s in self.shared_strings:
            escaped = _xml_escape(s)
            ss += f'<si><t>{escaped}</t></si>'
        ss += '</sst>'
        return ss

    def _build_sheet(self, sheet_idx):
        _, rows, data_validations, merge_cells = self.sheets[sheet_idx]
        s = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        s += '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
        # Column widths
        s += '<cols>'
        s += '<col min="1" max="1" width="32" customWidth="1"/>'
        s += '<col min="2" max="9" width="18" customWidth="1"/>'
        s += '</cols>'
        s += '<sheetData>'
        for r_idx, row in enumerate(rows):
            if not row:
                s += f'<row r="{r_idx + 1}"/>'
                continue
            s += f'<row r="{r_idx + 1}">'
            for c_idx, cell in enumerate(row):
                if cell is None:
                    continue
                # Determine cell type
                if len(cell) == 3:
                    value, style_id, cell_type = cell
                else:
                    value, style_id = cell
                    cell_type = 'string'
                if value is None:
                    continue
                ref = _cell_ref(r_idx, c_idx)
                if cell_type == 'formula':
                    # Formula cell: no t attribute
                    escaped_formula = _xml_escape(str(value))
                    s += f'<c r="{ref}" s="{style_id}"><f>{escaped_formula}</f><v>0</v></c>'
                elif cell_type == 'number':
                    # Number cell: no t attribute
                    s += f'<c r="{ref}" s="{style_id}"><v>{value}</v></c>'
                else:
                    # Shared string cell
                    si = self._get_ss_index(str(value))
                    s += f'<c r="{ref}" t="s" s="{style_id}"><v>{si}</v></c>'
            s += '</row>'
        s += '</sheetData>'
        # Merge cells
        if merge_cells:
            s += f'<mergeCells count="{len(merge_cells)}">'
            for mc in merge_cells:
                s += f'<mergeCell ref="{mc}"/>'
            s += '</mergeCells>'
        # Data validations
        if data_validations:
            s += f'<dataValidations count="{len(data_validations)}">'
            for sqref, formula1 in data_validations:
                s += f'<dataValidation type="list" allowBlank="1" sqref="{sqref}">'
                s += f'<formula1>{formula1}</formula1>'
                s += '</dataValidation>'
            s += '</dataValidations>'
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


def build_data_table(ages_data, benefit_key, plan_count, start_row):
    """Build a data table block for a specific benefit type.
    Returns (rows_list, num_data_rows) where rows_list contains the header + 172 data rows.
    start_row is 0-based row index where this table starts in the sheet.
    Ages 80-85 have no outpatient/dental data - use 0.
    """
    rows = []
    # Header row: Key, plan1, plan2, ...
    # (handled by caller who knows plan names)
    # Data rows: 172 rows (Pria_0..Pria_85, Wanita_0..Wanita_85)
    for gender in ["Pria", "Wanita"]:
        gender_key = gender.lower().replace("pria", "pria").replace("wanita", "wanita")
        for age in range(86):
            key = f"{gender}_{age}"
            entry = ages_data[age]
            gender_data = entry[gender_key]
            values = gender_data.get(benefit_key, [])
            row = [(key, 0)]  # Key column as string
            if values:
                for i in range(plan_count):
                    if i < len(values):
                        row.append((_num(values[i]), 5, 'number'))
                    else:
                        row.append((0, 5, 'number'))
            else:
                # No data for this age (e.g., outpatient for ages 80-85)
                for i in range(plan_count):
                    row.append((0, 5, 'number'))
            rows.append(row)
    return rows


def build_lookup_sheet(ages_data, sheet_type):
    """Build rows for a lookup sheet (Konvensional or Syariah).
    Returns (rows, data_validations, merge_cells).
    """
    is_syariah = (sheet_type == 'syariah')
    plans_smart = PLANS_SMART_SYARIAH if is_syariah else PLANS_SMART_KONV
    smart_count = len(plans_smart)
    regular_count = len(PLANS_REGULAR)

    title = "Tabel Premi MiUltimate HealthCare (Syariah)" if is_syariah else "Tabel Premi MiUltimate HealthCare (Konvensional)"

    rows = []
    # Row 0 (index 0): Header merged A1:I1
    header_row = [(HEADER_TEXT, 3)] + [None] * 8
    rows.append(header_row)
    # Row 1: Title
    rows.append([(title, 1)])
    # Row 2: Empty
    rows.append([])
    # Row 3 (index 3): Jenis Kelamin label + input
    rows.append([("Jenis Kelamin:", 1), ("Pria", 0)])
    # Row 4 (index 4): Usia label + input
    rows.append([("Usia:", 1), (0, 0, 'number')])
    # Row 5: Empty
    rows.append([])

    # Row 6 (index 6): Section header - Hospital
    rows.append([("Manfaat Perawatan Rumah Sakit", 2)])
    # Row 7: Plan names for regular
    row = [None] + [(p, 1) for p in PLANS_REGULAR]
    rows.append(row)

    # Now we need to know where data tables will start to write formulas.
    # Calculate data table positions:
    # Top section rows: 0-7 (8 rows) for header through plan names
    # Then formula rows for hospital regular (row 8)
    # Then smart plan names (row 9), smart formulas (row 10)
    # Empty (row 11)
    # Outpatient section header (row 12)
    # Outpatient plan names (row 13), formulas (row 14)
    # Outpatient smart plan names (row 15), smart formulas (row 16)
    # For syariah: Empty (row 17), Dental header (row 18), plan names (row 19),
    #   formulas (row 20), smart plan names (row 21), smart formulas (row 22)
    # Then empty row(s) before data tables

    if is_syariah:
        data_start_row = 25  # 0-based index where first data table starts
    else:
        data_start_row = 19  # 0-based index where first data table starts

    # Calculate data table block positions (0-based row indices)
    # Each block: 1 header row + 172 data rows = 173 rows
    block_size = 173  # header + 172 data rows

    hosp_reg_start = data_start_row
    hosp_smart_start = hosp_reg_start + block_size + 1  # +1 for gap
    outp_reg_start = hosp_smart_start + block_size + 1
    outp_smart_start = outp_reg_start + block_size + 1

    if is_syariah:
        dent_reg_start = outp_smart_start + block_size + 1
        dent_smart_start = dent_reg_start + block_size + 1

    # Helper to build INDEX/MATCH formula
    # Data starts at start + 1 (after header), key column is A, data columns start at B
    # Formula: =INDEX(COL_start:COL_end,MATCH($B$4&"_"&$B$5,$A$start:$A$end,0))
    def make_formula(block_start, plan_col_idx):
        """Make INDEX/MATCH formula.
        block_start: 0-based row of header.
        Data rows are block_start+1 to block_start+172 (0-based).
        In Excel 1-based: data from block_start+2 to block_start+173.
        plan_col_idx: 0-based column index in the data table (0=key col A, 1=first plan col B, etc.)
        """
        first_data_row_1based = block_start + 2  # 1-based row number
        last_data_row_1based = block_start + 173  # 1-based row number
        data_col = _col_letter(plan_col_idx)  # The plan data column
        key_col = "A"  # Key is always column A
        # =INDEX(B$27:B$198,MATCH($B$4&"_"&$B$5,$A$27:$A$198,0))
        formula = f'INDEX({data_col}${first_data_row_1based}:{data_col}${last_data_row_1based},MATCH($B$4&"_"&$B$5,${key_col}${first_data_row_1based}:${key_col}${last_data_row_1based},0))'
        return formula

    # Row 8 (index 8): Hospital Regular formulas
    formula_row = [None]
    for i in range(regular_count):
        f = make_formula(hosp_reg_start, i + 1)  # col B=1, C=2, etc.
        formula_row.append((f, 5, 'formula'))
    rows.append(formula_row)

    # Row 9: Smart plan names
    row = [None] + [(p, 1) for p in plans_smart]
    rows.append(row)

    # Row 10: Hospital Smart formulas
    formula_row = [None]
    for i in range(smart_count):
        f = make_formula(hosp_smart_start, i + 1)
        formula_row.append((f, 5, 'formula'))
    rows.append(formula_row)

    # Row 11: Empty
    rows.append([])

    # Row 12: Outpatient section header
    rows.append([("Manfaat Rawat Jalan", 2)])

    # Row 13: Outpatient plan names
    row = [None] + [(p, 1) for p in PLANS_REGULAR]
    rows.append(row)

    # Row 14: Outpatient Regular formulas
    formula_row = [None]
    for i in range(regular_count):
        f = make_formula(outp_reg_start, i + 1)
        formula_row.append((f, 5, 'formula'))
    rows.append(formula_row)

    # Row 15: Outpatient Smart plan names
    row = [None] + [(p, 1) for p in plans_smart]
    rows.append(row)

    # Row 16: Outpatient Smart formulas
    formula_row = [None]
    for i in range(smart_count):
        f = make_formula(outp_smart_start, i + 1)
        formula_row.append((f, 5, 'formula'))
    rows.append(formula_row)

    if is_syariah:
        # Row 17: Empty
        rows.append([])
        # Row 18: Dental section header
        rows.append([("Manfaat Perawatan Gigi", 2)])
        # Row 19: Dental plan names
        row = [None] + [(p, 1) for p in PLANS_REGULAR]
        rows.append(row)
        # Row 20: Dental Regular formulas
        formula_row = [None]
        for i in range(regular_count):
            f = make_formula(dent_reg_start, i + 1)
            formula_row.append((f, 5, 'formula'))
        rows.append(formula_row)
        # Row 21: Dental Smart plan names
        row = [None] + [(p, 1) for p in plans_smart]
        rows.append(row)
        # Row 22: Dental Smart formulas
        formula_row = [None]
        for i in range(smart_count):
            f = make_formula(dent_smart_start, i + 1)
            formula_row.append((f, 5, 'formula'))
        rows.append(formula_row)

    # Pad with empty rows until data_start_row
    while len(rows) < data_start_row:
        rows.append([])

    # Now build data tables
    # Block 1: Hospital Regular
    header = [("Key", 1)] + [(p, 1) for p in PLANS_REGULAR]
    rows.append(header)
    data_rows = build_data_table(ages_data, 'hospital', regular_count, hosp_reg_start)
    rows.extend(data_rows)

    # Gap
    rows.append([])

    # Block 2: Hospital Smart
    header = [("Key", 1)] + [(p, 1) for p in plans_smart]
    rows.append(header)
    data_rows = build_data_table(ages_data, 'hospital_smart', smart_count, hosp_smart_start)
    rows.extend(data_rows)

    # Gap
    rows.append([])

    # Block 3: Outpatient Regular
    header = [("Key", 1)] + [(p, 1) for p in PLANS_REGULAR]
    rows.append(header)
    data_rows = build_data_table(ages_data, 'outpatient', regular_count, outp_reg_start)
    rows.extend(data_rows)

    # Gap
    rows.append([])

    # Block 4: Outpatient Smart
    header = [("Key", 1)] + [(p, 1) for p in plans_smart]
    rows.append(header)
    data_rows = build_data_table(ages_data, 'outpatient_smart', smart_count, outp_smart_start)
    rows.extend(data_rows)

    if is_syariah:
        # Gap
        rows.append([])

        # Block 5: Dental Regular
        header = [("Key", 1)] + [(p, 1) for p in PLANS_REGULAR]
        rows.append(header)
        data_rows = build_data_table(ages_data, 'dental', regular_count, dent_reg_start)
        rows.extend(data_rows)

        # Gap
        rows.append([])

        # Block 6: Dental Smart
        header = [("Key", 1)] + [(p, 1) for p in plans_smart]
        rows.append(header)
        data_rows = build_data_table(ages_data, 'dental_smart', smart_count, dent_smart_start)
        rows.extend(data_rows)

    # Data validation for gender dropdown in B4 (0-based row 3)
    data_validations = [("B4", '"Pria,Wanita"')]
    # Merge cells for header
    merge_cells = ["A1:I1"]

    return rows, data_validations, merge_cells


def main():
    print("Parsing Konvensional PDF...")
    konv_data = parse_konvensional()
    print(f"  Extracted data for {len(konv_data)} ages")

    print("Parsing Syariah PDF...")
    syariah_data = parse_syariah()
    print(f"  Extracted data for {len(syariah_data)} ages")

    print("Building Excel file...")
    writer = XlsxWriter()

    konv_rows, konv_dv, konv_mc = build_lookup_sheet(konv_data, 'konvensional')
    syariah_rows, syariah_dv, syariah_mc = build_lookup_sheet(syariah_data, 'syariah')

    writer.add_sheet("Konvensional", konv_rows, konv_dv, konv_mc)
    writer.add_sheet("Syariah", syariah_rows, syariah_dv, syariah_mc)

    # Add hyperlinks for row 0 col 0 on both sheets
    writer.add_hyperlink(0, 0, 0, WA_LINK)
    writer.add_hyperlink(1, 0, 0, WA_LINK)

    writer.save(OUTPUT_XLSX)
    print(f"Excel file saved to: {OUTPUT_XLSX}")

    # Verify
    with zipfile.ZipFile(OUTPUT_XLSX) as z:
        print(f"  Valid xlsx with {len(z.namelist())} files")
        print(f"  Contents: {z.namelist()}")


if __name__ == "__main__":
    main()
