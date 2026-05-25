#!/usr/bin/env python3
"""Generate Kalkulator_Premi_MiUHC.xlsx - MiUHC Plans 1-5 Premium Calculator."""
import zipfile
import os

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
OUTPUT_XLSX = os.path.join(SCRIPT_DIR, 'Kalkulator_Premi_MiUHC.xlsx')

WA_LINK = "https://wa.me/6287781896087"
HEADER_TEXT = "PETRUS JAKUB MANULIFE 087781896087"

# Premium data: PREMIUM_DATA[plan][gender][age_band] = annual premium (IDR)
PREMIUM_DATA = {
    "Plan 1": {
        "Pria":   {"20-30": 3500000, "31-40": 4200000, "41-50": 5800000, "51-60": 8500000, "61-70": 12000000},
        "Wanita": {"20-30": 3800000, "31-40": 4500000, "41-50": 6200000, "51-60": 9000000, "61-70": 12800000},
    },
    "Plan 2": {
        "Pria":   {"20-30": 5200000, "31-40": 6300000, "41-50": 8700000, "51-60": 12500000, "61-70": 17500000},
        "Wanita": {"20-30": 5600000, "31-40": 6800000, "41-50": 9300000, "51-60": 13200000, "61-70": 18500000},
    },
    "Plan 3": {
        "Pria":   {"20-30": 7800000, "31-40": 9500000, "41-50": 13000000, "51-60": 18800000, "61-70": 26000000},
        "Wanita": {"20-30": 8400000, "31-40": 10200000, "41-50": 14000000, "51-60": 20000000, "61-70": 27500000},
    },
    "Plan 4": {
        "Pria":   {"20-30": 11500000, "31-40": 14000000, "41-50": 19500000, "51-60": 28000000, "61-70": 39000000},
        "Wanita": {"20-30": 12500000, "31-40": 15200000, "41-50": 21000000, "51-60": 30000000, "61-70": 41500000},
    },
    "Plan 5": {
        "Pria":   {"20-30": 16800000, "31-40": 20500000, "41-50": 28500000, "51-60": 41000000, "61-70": 57000000},
        "Wanita": {"20-30": 18200000, "31-40": 22000000, "41-50": 30500000, "51-60": 43500000, "61-70": 60000000},
    },
}

AGE_BANDS = ["20-30", "31-40", "41-50", "51-60", "61-70"]
GENDERS = ["Pria", "Wanita"]
PLANS = ["Plan 1", "Plan 2", "Plan 3", "Plan 4", "Plan 5"]


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
        data_validations: list of tuples:
          For list type: (sqref, "list", formula1, None)
          For whole number: (sqref, "whole", formula1, formula2)
        merge_cells: list of ref strings e.g. ["A1:F1"]
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
            wb += f'<sheet name="{_xml_escape(name)}" sheetId="{i+1}" r:id="rId{i+1}"/>'
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
        # 5 = number format #,##0
        # 6 = input cell (yellow background)
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
        s += '<fills count="4">'
        s += '<fill><patternFill patternType="none"/></fill>'  # 0: none
        s += '<fill><patternFill patternType="gray125"/></fill>'  # 1: gray125
        s += '<fill><patternFill patternType="solid"><fgColor rgb="FF00B050"/></patternFill></fill>'  # 2: green
        s += '<fill><patternFill patternType="solid"><fgColor rgb="FFFFFF00"/></patternFill></fill>'  # 3: yellow
        s += '</fills>'
        # Borders
        s += '<borders count="2">'
        s += '<border><left/><right/><top/><bottom/><diagonal/></border>'  # 0: no border
        s += '<border>'  # 1: thin border all sides
        s += '<left style="thin"><color auto="1"/></left>'
        s += '<right style="thin"><color auto="1"/></right>'
        s += '<top style="thin"><color auto="1"/></top>'
        s += '<bottom style="thin"><color auto="1"/></bottom>'
        s += '<diagonal/>'
        s += '</border>'
        s += '</borders>'
        # CellStyleXfs
        s += '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>'
        # CellXfs
        s += '<cellXfs count="7">'
        s += '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>'  # 0: default
        s += '<xf numFmtId="0" fontId="1" fillId="0" borderId="0" xfId="0" applyFont="1"/>'  # 1: bold
        s += '<xf numFmtId="0" fontId="2" fillId="2" borderId="0" xfId="0" applyFont="1" applyFill="1"/>'  # 2: green bg white bold
        s += '<xf numFmtId="0" fontId="3" fillId="0" borderId="0" xfId="0" applyFont="1"/>'  # 3: bold green text
        s += '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyAlignment="1"><alignment horizontal="right"/></xf>'  # 4: right-aligned
        s += '<xf numFmtId="164" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>'  # 5: number format #,##0
        s += '<xf numFmtId="0" fontId="0" fillId="3" borderId="1" xfId="0" applyFill="1" applyBorder="1"/>'  # 6: input cell (yellow bg + border)
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
        s += '<col min="1" max="1" width="22" customWidth="1"/>'
        s += '<col min="2" max="2" width="24" customWidth="1"/>'
        s += '<col min="3" max="7" width="16" customWidth="1"/>'
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
                    escaped_formula = _xml_escape(str(value))
                    s += f'<c r="{ref}" s="{style_id}"><f>{escaped_formula}</f><v>0</v></c>'
                elif cell_type == 'number':
                    s += f'<c r="{ref}" s="{style_id}"><v>{value}</v></c>'
                else:
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
            for dv in data_validations:
                sqref, dv_type, formula1, formula2 = dv
                if dv_type == "list":
                    s += f'<dataValidation type="list" allowBlank="1" sqref="{sqref}">'
                    s += f'<formula1>{_xml_escape(formula1)}</formula1>'
                    s += '</dataValidation>'
                elif dv_type == "whole":
                    s += f'<dataValidation type="whole" operator="between" allowBlank="1" sqref="{sqref}">'
                    s += f'<formula1>{formula1}</formula1>'
                    s += f'<formula2>{formula2}</formula2>'
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


def build_calculator_sheet(sheet_type):
    """Build rows for a calculator sheet (Konvensional or Syariah).

    Layout:
      Row 1: Header (merged A1:F1, hyperlinked)
      Row 2: Subtitle
      Row 3: Empty
      Row 4: Nama Tertanggung: | [input]
      Row 5: Jenis Kelamin:    | [dropdown]
      Row 6: Tanggal Lahir:    | [input]
      Row 7: Usia:             | [formula]
      Row 8: Plan:             | [dropdown]
      Row 9: Empty
      Row 10: HASIL PERHITUNGAN
      Row 11: Premi Tahunan (Rp): | [lookup formula]
      Row 12: Premi Bulanan (Rp): | [formula =B11/10]
      Row 13: Empty
      Row 14: Empty
      Row 15: (Data table header)
      Row 16-25: (Data table rows)

    Returns (rows, data_validations, merge_cells).
    """
    if sheet_type == 'konvensional':
        subtitle = "MiUHC Konvensional - Efektif 1 September 2025"
    else:
        subtitle = "MiUHC Syariah - Efektif 1 Januari 2024"

    rows = []

    # Row 0 (Excel row 1): Header merged A1:F1
    header_row = [(HEADER_TEXT, 3)] + [None] * 5
    rows.append(header_row)

    # Row 1 (Excel row 2): Subtitle
    rows.append([(subtitle, 1)])

    # Row 2 (Excel row 3): Empty
    rows.append([])

    # Row 3 (Excel row 4): Nama Tertanggung
    rows.append([("Nama Tertanggung:", 1), ("", 6)])

    # Row 4 (Excel row 5): Jenis Kelamin with dropdown
    rows.append([("Jenis Kelamin:", 1), ("Pria", 6)])

    # Row 5 (Excel row 6): Tanggal Lahir
    rows.append([("Tanggal Lahir:", 1), ("", 6)])

    # Row 6 (Excel row 7): Usia - formula calculating age from B6
    # Using INT((TODAY()-B6)/365.25) which works well for date serial numbers
    age_formula = 'INT((TODAY()-B6)/365.25)'
    rows.append([("Usia:", 1), (age_formula, 0, 'formula')])

    # Row 7 (Excel row 8): Plan with dropdown
    rows.append([("Plan:", 1), ("Plan 1", 6)])

    # Row 8 (Excel row 9): Empty
    rows.append([])

    # Row 9 (Excel row 10): HASIL PERHITUNGAN section header
    rows.append([("HASIL PERHITUNGAN", 2)])

    # Row 10 (Excel row 11): Premi Tahunan with lookup formula
    # The data table starts at row 16 (0-based row 15).
    # Key column is A16:A25, Plan headers are B15:F15, Data is B16:F25.
    # Age band formula: IF(B7<=30,"20-30",IF(B7<=40,"31-40",IF(B7<=50,"41-50",IF(B7<=60,"51-60","61-70"))))
    # Lookup key: B5&"_"&age_band
    # Formula: IFERROR(INDEX(B16:F25, MATCH(B5&"_"&age_band, A16:A25, 0), MATCH(B8, B15:F15, 0)), "")
    age_band_part = 'IF(B7<=30,"20-30",IF(B7<=40,"31-40",IF(B7<=50,"41-50",IF(B7<=60,"51-60","61-70"))))'
    lookup_key = f'B5&"_"&{age_band_part}'
    premi_formula = f'IFERROR(INDEX(B16:F25,MATCH({lookup_key},A16:A25,0),MATCH(B8,B15:F15,0)),"")'
    rows.append([("Premi Tahunan (Rp):", 1), (premi_formula, 5, 'formula')])

    # Row 11 (Excel row 12): Premi Bulanan = Tahunan / 10
    rows.append([("Premi Bulanan (Rp):", 1), ('IFERROR(B11/10,"")', 5, 'formula')])

    # Row 12 (Excel row 13): Empty
    rows.append([])

    # Row 13 (Excel row 14): Empty
    rows.append([])

    # Row 14 (Excel row 15): Data table header
    # Columns: A=Key, B=Plan 1, C=Plan 2, D=Plan 3, E=Plan 4, F=Plan 5
    data_header = [("Key", 1)]
    for plan in PLANS:
        data_header.append((plan, 1))
    rows.append(data_header)

    # Rows 15-24 (Excel rows 16-25): Data table rows
    # Order: Pria_20-30, Pria_31-40, ..., Pria_61-70, Wanita_20-30, ..., Wanita_61-70
    for gender in GENDERS:
        for age_band in AGE_BANDS:
            key = f"{gender}_{age_band}"
            row = [(key, 0)]
            for plan in PLANS:
                row.append((PREMIUM_DATA[plan][gender][age_band], 5, 'number'))
            rows.append(row)

    # Data validations
    data_validations = [
        ("B5", "list", '"Pria,Wanita"', None),
        ("B8", "list", '"Plan 1,Plan 2,Plan 3,Plan 4,Plan 5"', None),
    ]

    # Merge cells for header
    merge_cells = ["A1:F1"]

    return rows, data_validations, merge_cells


def main():
    print("Building MiUHC Plans 1-5 Premium Calculator...")
    writer = XlsxWriter()

    # Build Konvensional sheet
    konv_rows, konv_dv, konv_mc = build_calculator_sheet('konvensional')
    writer.add_sheet("Konvensional", konv_rows, konv_dv, konv_mc)

    # Build Syariah sheet
    syariah_rows, syariah_dv, syariah_mc = build_calculator_sheet('syariah')
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
