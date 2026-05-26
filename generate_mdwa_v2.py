#!/usr/bin/env python3
"""Generate MDWA_Kalkulator_v2.xlsx - FIXED version with proper dropdown visibility.

Key fixes from v1:
- showDropDown="0" in dataValidation XML (counterintuitive: "0" = SHOW the dropdown arrow)
- Explicit yellow fill (FFFFFF00) on input cells with blue border
- Proper data validation XML structure with showInputMessage and showErrorMessage

Each Plan (A, B, C) has its own sheet with:
- Visible YELLOW input cells with dropdown menus for Gender, Age, Premium, etc.
- Formulas that auto-calculate benefits based on inputs
- Lookup tables in hidden rows (50+) for INDEX/MATCH

Sheets: PLAN A, PLAN B, PLAN C, KOMISI, PERBANDINGAN, RINGKASAN
"""
import zipfile
import os
import re

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
OUTPUT_XLSX = os.path.join(SCRIPT_DIR, 'MDWA_Kalkulator_v2.xlsx')


def xml_escape(s):
    return s.replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;').replace('"', '&quot;')


def col_letter(col_idx):
    """Convert 0-based column index to Excel letter (0=A, 1=B, etc.)."""
    result = ''
    idx = col_idx
    while True:
        result = chr(65 + idx % 26) + result
        idx = idx // 26 - 1
        if idx < 0:
            break
    return result


class SheetBuilder:
    """Helper to build worksheet XML with formulas, data validation, merges."""

    def __init__(self, ss_map):
        self.ss_map = ss_map
        self.rows = []
        self.merges = []
        self.col_widths = {}
        self.data_validations = []

    def set_col_width(self, col, width):
        self.col_widths[col] = width

    def add_row(self, row_num, cells, height=None):
        """cells: list of (col_idx, value, style, cell_type)
        cell_type: 's'=string, 'n'=number, 'f'=formula, None=auto
        """
        self.rows.append((row_num, cells, height))

    def add_merge(self, ref):
        self.merges.append(ref)

    def add_data_validation(self, sqref, formula1, allow_blank=True):
        self.data_validations.append((sqref, formula1, allow_blank))

    def _cell_xml(self, col_idx, row_num, value, style=0, cell_type=None):
        ref = col_letter(col_idx) + str(row_num)
        if cell_type == 'f':
            return '<c r="' + ref + '" s="' + str(style) + '"><f>' + xml_escape(str(value)) + '</f><v>0</v></c>'
        elif cell_type == 'n' or (cell_type is None and isinstance(value, (int, float))):
            return '<c r="' + ref + '" s="' + str(style) + '"><v>' + str(value) + '</v></c>'
        elif isinstance(value, str):
            if value in self.ss_map:
                si = self.ss_map[value]
                return '<c r="' + ref + '" s="' + str(style) + '" t="s"><v>' + str(si) + '</v></c>'
            else:
                return '<c r="' + ref + '" s="' + str(style) + '" t="inlineStr"><is><t>' + xml_escape(value) + '</t></is></c>'
        else:
            return '<c r="' + ref + '" s="' + str(style) + '"/>'

    def build(self):
        s = chr(60) + '?xml version="1.0" encoding="UTF-8" standalone="yes"?' + chr(62)
        s += '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
        s += '<sheetViews><sheetView workbookViewId="0"/></sheetViews>'
        s += '<sheetFormatPr defaultColWidth="14.0" defaultRowHeight="15.0"/>'
        if self.col_widths:
            s += '<cols>'
            for col, w in sorted(self.col_widths.items()):
                s += '<col customWidth="1" min="' + str(col) + '" max="' + str(col) + '" width="' + str(w) + '"/>'
            s += '</cols>'
        s += '<sheetData>'
        for row_num, cells, height in self.rows:
            ht_attr = ' ht="' + str(height) + '" customHeight="1"' if height else ''
            s += '<row r="' + str(row_num) + '"' + ht_attr + '>'
            for cell_data in cells:
                if len(cell_data) == 3:
                    col_idx, value, style = cell_data
                    s += self._cell_xml(col_idx, row_num, value, style)
                elif len(cell_data) == 4:
                    col_idx, value, style, ctype = cell_data
                    s += self._cell_xml(col_idx, row_num, value, style, ctype)
            s += '</row>'
        s += '</sheetData>'
        if self.merges:
            s += '<mergeCells count="' + str(len(self.merges)) + '">'
            for m in self.merges:
                s += '<mergeCell ref="' + m + '"/>'
            s += '</mergeCells>'
        if self.data_validations:
            s += '<dataValidations count="' + str(len(self.data_validations)) + '">'
            for sqref, formula1, allow_blank in self.data_validations:
                blank = ' allowBlank="1"' if allow_blank else ''
                s += '<dataValidation type="list" showDropDown="0" sqref="' + sqref + '"' + blank + ' showInputMessage="1" showErrorMessage="1">'
                s += '<formula1>' + xml_escape(formula1) + '</formula1>'
                s += '</dataValidation>'
            s += '</dataValidations>'
        s += '</worksheet>'
        return s


# ==============================================================
# STYLES
# ==============================================================
# Style indices:
# 0: default
# 1: header green bg white bold text centered
# 2: data cell with border centered
# 3: title bold green font large
# 4: data cell light green bg with border
# 5: section header dark green bg white bold large
# 6: number format with comma, border, centered
# 7: bold black left aligned
# 8: input cell (YELLOW bg, blue border, bold blue font) - PROMINENT
# 9: result cell (light blue bg, border, number format)
# 10: label bold black with border
# 11: note italic gray
# 12: percentage format with border centered

def build_styles_xml():
    s = chr(60) + '?xml version="1.0" encoding="UTF-8" standalone="yes"?' + chr(62)
    s += '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
    s += '<numFmts count="2">'
    s += '<numFmt numFmtId="164" formatCode="#,##0"/>'
    s += '<numFmt numFmtId="165" formatCode="0.00%"/>'
    s += '</numFmts>'
    # fonts: 0=default, 1=bold white, 2=bold green title, 3=bold black, 4=normal small,
    # 5=bold white large, 6=bold blue (input), 7=italic gray (note)
    s += '<fonts count="8">'
    s += '<font><sz val="10.0"/><color rgb="FF000000"/><name val="Arial"/></font>'
    s += '<font><b/><sz val="10.0"/><color rgb="FFFFFFFF"/><name val="Arial"/></font>'
    s += '<font><b/><sz val="14.0"/><color rgb="FF1E7145"/><name val="Arial"/></font>'
    s += '<font><b/><sz val="11.0"/><color rgb="FF000000"/><name val="Arial"/></font>'
    s += '<font><sz val="9.0"/><color rgb="FF333333"/><name val="Arial"/></font>'
    s += '<font><b/><sz val="12.0"/><color rgb="FFFFFFFF"/><name val="Arial"/></font>'
    s += '<font><b/><sz val="11.0"/><color rgb="FF0000FF"/><name val="Arial"/></font>'
    s += '<font><i/><sz val="9.0"/><color rgb="FF666666"/><name val="Arial"/></font>'
    s += '</fonts>'
    # fills: 0=none, 1=gray, 2=manulife green, 3=light green, 4=white, 5=dark green,
    # 6=YELLOW (input highlight), 7=light blue (result)
    s += '<fills count="8">'
    s += '<fill><patternFill patternType="none"/></fill>'
    s += '<fill><patternFill patternType="lightGray"/></fill>'
    s += '<fill><patternFill patternType="solid"><fgColor rgb="FF1E7145"/><bgColor rgb="FF1E7145"/></patternFill></fill>'
    s += '<fill><patternFill patternType="solid"><fgColor rgb="FFE8F5EE"/><bgColor rgb="FFE8F5EE"/></patternFill></fill>'
    s += '<fill><patternFill patternType="solid"><fgColor rgb="FFFFFFFF"/><bgColor rgb="FFFFFFFF"/></patternFill></fill>'
    s += '<fill><patternFill patternType="solid"><fgColor rgb="FF00B050"/><bgColor rgb="FF00B050"/></patternFill></fill>'
    s += '<fill><patternFill patternType="solid"><fgColor rgb="FFFFFF00"/><bgColor rgb="FFFFFF00"/></patternFill></fill>'
    s += '<fill><patternFill patternType="solid"><fgColor rgb="FFE6F3FF"/><bgColor rgb="FFE6F3FF"/></patternFill></fill>'
    s += '</fills>'
    # borders
    s += '<borders count="3">'
    s += '<border/>'
    s += '<border><left style="thin"><color rgb="FFCCCCCC"/></left>'
    s += '<right style="thin"><color rgb="FFCCCCCC"/></right>'
    s += '<top style="thin"><color rgb="FFCCCCCC"/></top>'
    s += '<bottom style="thin"><color rgb="FFCCCCCC"/></bottom></border>'
    s += '<border><left style="medium"><color rgb="FF0000FF"/></left>'
    s += '<right style="medium"><color rgb="FF0000FF"/></right>'
    s += '<top style="medium"><color rgb="FF0000FF"/></top>'
    s += '<bottom style="medium"><color rgb="FF0000FF"/></bottom></border>'
    s += '</borders>'
    s += '<cellStyleXfs count="1">'
    s += '<xf borderId="0" fillId="0" fontId="0" numFmtId="0"/>'
    s += '</cellStyleXfs>'
    # cellXfs (styles 0-12)
    s += '<cellXfs count="13">'
    # 0: default
    s += '<xf borderId="0" fillId="0" fontId="0" numFmtId="0" xfId="0" applyAlignment="1"><alignment vertical="center" wrapText="1"/></xf>'
    # 1: header green bg white bold centered
    s += '<xf borderId="1" fillId="2" fontId="1" numFmtId="0" xfId="0" applyAlignment="1" applyBorder="1" applyFill="1" applyFont="1"><alignment horizontal="center" vertical="center" wrapText="1"/></xf>'
    # 2: data cell bordered centered
    s += '<xf borderId="1" fillId="0" fontId="0" numFmtId="0" xfId="0" applyAlignment="1" applyBorder="1"><alignment horizontal="center" vertical="center" wrapText="1"/></xf>'
    # 3: title bold green font
    s += '<xf borderId="0" fillId="0" fontId="2" numFmtId="0" xfId="0" applyAlignment="1" applyFont="1"><alignment vertical="center" wrapText="1"/></xf>'
    # 4: data cell light green bg
    s += '<xf borderId="1" fillId="3" fontId="0" numFmtId="0" xfId="0" applyAlignment="1" applyBorder="1" applyFill="1"><alignment horizontal="center" vertical="center" wrapText="1"/></xf>'
    # 5: section header dark green bg white bold large
    s += '<xf borderId="1" fillId="5" fontId="5" numFmtId="0" xfId="0" applyAlignment="1" applyBorder="1" applyFill="1" applyFont="1"><alignment horizontal="center" vertical="center" wrapText="1"/></xf>'
    # 6: number format with comma, border, centered
    s += '<xf borderId="1" fillId="0" fontId="0" numFmtId="164" xfId="0" applyAlignment="1" applyBorder="1" applyNumberFormat="1"><alignment horizontal="center" vertical="center" wrapText="1"/></xf>'
    # 7: bold black left aligned
    s += '<xf borderId="0" fillId="0" fontId="3" numFmtId="0" xfId="0" applyAlignment="1" applyFont="1"><alignment vertical="center" wrapText="1"/></xf>'
    # 8: input cell (YELLOW bg, blue border, bold blue font) - VERY VISIBLE
    s += '<xf borderId="2" fillId="6" fontId="6" numFmtId="0" xfId="0" applyAlignment="1" applyBorder="1" applyFill="1" applyFont="1"><alignment horizontal="center" vertical="center" wrapText="1"/></xf>'
    # 9: result cell (light blue bg, border, number format)
    s += '<xf borderId="1" fillId="7" fontId="3" numFmtId="164" xfId="0" applyAlignment="1" applyBorder="1" applyFill="1" applyFont="1" applyNumberFormat="1"><alignment horizontal="center" vertical="center" wrapText="1"/></xf>'
    # 10: label bold black with border
    s += '<xf borderId="1" fillId="0" fontId="3" numFmtId="0" xfId="0" applyAlignment="1" applyBorder="1" applyFont="1"><alignment vertical="center" wrapText="1"/></xf>'
    # 11: note italic gray
    s += '<xf borderId="0" fillId="0" fontId="7" numFmtId="0" xfId="0" applyAlignment="1" applyFont="1"><alignment vertical="center" wrapText="1"/></xf>'
    # 12: result cell text (light blue bg, border, no number format)
    s += '<xf borderId="1" fillId="7" fontId="3" numFmtId="0" xfId="0" applyAlignment="1" applyBorder="1" applyFill="1" applyFont="1"><alignment horizontal="center" vertical="center" wrapText="1"/></xf>'
    s += '</cellXfs>'
    s += '<cellStyles count="1"><cellStyle xfId="0" name="Normal" builtinId="0"/></cellStyles>'
    s += '<dxfs count="0"/>'
    s += '</styleSheet>'
    return s


def build_theme_xml():
    t = chr(60) + '?xml version="1.0" encoding="UTF-8" standalone="yes"?' + chr(62)
    t += '<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="MDWA">'
    t += '<a:themeElements>'
    t += '<a:clrScheme name="MDWA">'
    t += '<a:dk1><a:srgbClr val="000000"/></a:dk1>'
    t += '<a:lt1><a:srgbClr val="FFFFFF"/></a:lt1>'
    t += '<a:dk2><a:srgbClr val="000000"/></a:dk2>'
    t += '<a:lt2><a:srgbClr val="FFFFFF"/></a:lt2>'
    t += '<a:accent1><a:srgbClr val="1E7145"/></a:accent1>'
    t += '<a:accent2><a:srgbClr val="00B050"/></a:accent2>'
    t += '<a:accent3><a:srgbClr val="9BBB59"/></a:accent3>'
    t += '<a:accent4><a:srgbClr val="8064A2"/></a:accent4>'
    t += '<a:accent5><a:srgbClr val="4BACC6"/></a:accent5>'
    t += '<a:accent6><a:srgbClr val="F79646"/></a:accent6>'
    t += '<a:hlink><a:srgbClr val="0000FF"/></a:hlink>'
    t += '<a:folHlink><a:srgbClr val="0000FF"/></a:folHlink>'
    t += '</a:clrScheme>'
    t += '<a:fontScheme name="MDWA">'
    t += '<a:majorFont><a:latin typeface="Arial"/><a:ea typeface="Arial"/><a:cs typeface="Arial"/></a:majorFont>'
    t += '<a:minorFont><a:latin typeface="Arial"/><a:ea typeface="Arial"/><a:cs typeface="Arial"/></a:minorFont>'
    t += '</a:fontScheme>'
    t += '<a:fmtScheme name="Office">'
    t += '<a:fillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:fillStyleLst>'
    t += '<a:lnStyleLst><a:ln><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln><a:ln><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln><a:ln><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln></a:lnStyleLst>'
    t += '<a:effectStyleLst><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle></a:effectStyleLst>'
    t += '<a:bgFillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:bgFillStyleLst>'
    t += '</a:fmtScheme>'
    t += '</a:themeElements></a:theme>'
    return t


def build_shared_strings(strings):
    ss = chr(60) + '?xml version="1.0" encoding="UTF-8" standalone="yes"?' + chr(62)
    ss += '<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"'
    ss += ' count="' + str(len(strings)) + '" uniqueCount="' + str(len(strings)) + '">'
    for s in strings:
        escaped = xml_escape(s)
        ss += '<si><t>' + escaped + '</t></si>'
    ss += '</sst>'
    return ss


# ==============================================================
# PLAN A SHEET - ANUITAS
# ==============================================================
def build_plan_a_sheet(ss_map):
    """PLAN A sheet with own Gender/Age/Premium/MasaBayar/Coverage inputs.
    
    Input cells (YELLOW background):
    - C4: Jenis Kelamin (Pria/Wanita) dropdown
    - C5: Usia Masuk (number)
    - C6: Premi Tahunan IDR (number)
    - C7: Masa Bayar (Single/5 Tahun/10 Tahun) dropdown
    - C8: Masa Pertanggungan (20 Tahun/30 Tahun) dropdown
    
    Results (row 10+):
    - Status Kelayakan
    - Level Banding
    - Manfaat Tahapan Tahunan (%)
    - Manfaat Tahapan Tahunan (nominal)
    - Mulai Pembayaran
    - Jumlah Tahun Pembayaran
    - Total Tahapan
    - ADB & TPD per Tahun
    - Total Seluruh Manfaat
    
    Lookup table in rows 50+ for Banding 2 extra percentages.
    """
    sb = SheetBuilder(ss_map)
    sb.set_col_width(1, 8)   # A - spacer
    sb.set_col_width(2, 32)  # B - labels
    sb.set_col_width(3, 28)  # C - inputs/values
    sb.set_col_width(4, 30)  # D - notes
    sb.set_col_width(5, 20)  # E

    # Row 1-2: Title
    sb.add_row(1, [(1, 'PLAN A - ANUITAS', 3)], height=28)
    sb.add_merge('B1:D1')
    sb.add_row(2, [(1, 'Manfaat Tahapan Tahunan (Annual Payout)', 11)])

    # Row 3: Section header INPUT
    sb.add_row(3, [(1, 'INPUT DATA', 5), (2, '', 5), (3, '', 5)])
    sb.add_merge('B3:D3')

    # Row 4: Gender
    sb.add_row(4, [(1, 'JENIS KELAMIN:', 10), (2, 'Pria', 8)])
    sb.add_data_validation('C4', '"Pria,Wanita"')

    # Row 5: Age
    sb.add_row(5, [(1, 'USIA MASUK:', 10), (2, 30, 8, 'n')])

    # Row 6: Premium
    sb.add_row(6, [(1, 'PREMI TAHUNAN (IDR):', 10), (2, 100000000, 8, 'n')])

    # Row 7: Payment Term
    sb.add_row(7, [(1, 'MASA BAYAR:', 10), (2, 'Single', 8)])
    sb.add_data_validation('C7', '"Single,5 Tahun,10 Tahun"')

    # Row 8: Coverage
    sb.add_row(8, [(1, 'MASA PERTANGGUNGAN:', 10), (2, '20 Tahun', 8)])
    sb.add_data_validation('C8', '"20 Tahun,30 Tahun"')

    # Row 9: blank separator
    sb.add_row(9, [])

    # Row 10: HASIL section header
    sb.add_row(10, [(1, 'HASIL PERHITUNGAN', 5), (2, '', 5), (3, '', 5)])
    sb.add_merge('B10:D10')

    # Row 11: Status Kelayakan
    # Single: max 85, Regular: max 70
    elig_f = 'IF(C7="Single",IF(C5<=85,"ELIGIBLE","TIDAK ELIGIBLE - Usia maks 85"),IF(C5<=70,"ELIGIBLE","TIDAK ELIGIBLE - Usia maks 70"))'
    sb.add_row(11, [(1, 'STATUS KELAYAKAN:', 10), (2, elig_f, 12, 'f')])

    # Row 12: Level Banding
    # Single: >=500M = B2, else B1; Regular: >=100M = B2, else B1
    banding_f = 'IF(C7="Single",IF(C6>=500000000,"Banding 2","Banding 1"),IF(C6>=100000000,"Banding 2","Banding 1"))'
    sb.add_row(12, [(1, 'LEVEL BANDING:', 10), (2, banding_f, 12, 'f')])

    # Row 13: Tahapan % (Banding 1)
    # Plan A: 20yr: Single=9.5%, 5yr=65%, 10yr=235%; 30yr: Single=7%, 5yr=40%, 10yr=105%
    tahapan_pct_f = 'IF(C8="20 Tahun",IF(C7="Single",9.5,IF(C7="5 Tahun",65,235)),IF(C7="Single",7,IF(C7="5 Tahun",40,105)))'
    sb.add_row(13, [(1, 'MANFAAT TAHAPAN TAHUNAN (%):', 10), (2, tahapan_pct_f, 9, 'f'), (3, '% dari Premi (Banding 1)', 11)])

    # Row 14: Banding 2 Extra %
    # 20yr: Single+6%, 5yr+30%, 10yr+60%; 30yr: Single+9%, 5yr+45%, 10yr+90%
    extra_pct_f = 'IF(C12="Banding 2",IF(C8="20 Tahun",IF(C7="Single",6,IF(C7="5 Tahun",30,60)),IF(C7="Single",9,IF(C7="5 Tahun",45,90))),0)'
    sb.add_row(14, [(1, 'BANDING 2 EXTRA (%):', 10), (2, extra_pct_f, 9, 'f'), (3, 'Tambahan jika Banding 2', 11)])

    # Row 15: Total Tahapan %
    total_pct_f = 'C13+C14'
    sb.add_row(15, [(1, 'TOTAL TAHAPAN % PER TAHUN:', 10), (2, total_pct_f, 9, 'f')])

    # Row 16: Nominal Tahapan per tahun
    tahapan_nom_f = 'C15*C6/100'
    sb.add_row(16, [(1, 'MANFAAT TAHAPAN TAHUNAN (IDR):', 10), (2, tahapan_nom_f, 9, 'f'), (3, 'Per tahun', 11)])

    # Row 17: Mulai Pembayaran
    # Single: yr 6, 5yr: yr 10, 10yr: yr 15
    mulai_f = 'IF(C7="Single","Tahun ke-6",IF(C7="5 Tahun","Tahun ke-10","Tahun ke-15"))'
    sb.add_row(17, [(1, 'MULAI PEMBAYARAN:', 10), (2, mulai_f, 12, 'f')])

    # Row 18: Jumlah tahun pembayaran
    # coverage_years - start_year + 1
    # 20yr: Single=15, 5yr=11, 10yr=6; 30yr: Single=25, 5yr=21, 10yr=16
    jml_f = 'IF(C8="20 Tahun",IF(C7="Single",15,IF(C7="5 Tahun",11,6)),IF(C7="Single",25,IF(C7="5 Tahun",21,16)))'
    sb.add_row(18, [(1, 'JUMLAH TAHUN PEMBAYARAN:', 10), (2, jml_f, 9, 'f'), (3, 'tahun', 11)])

    # Row 19: Total Tahapan (seluruh periode)
    total_tahapan_f = 'C16*C18'
    sb.add_row(19, [(1, 'TOTAL TAHAPAN (SELURUH PERIODE):', 10), (2, total_tahapan_f, 9, 'f')])

    # Row 20: ADB & TPD per tahun
    # Single: 25% * premi / masa_pertanggungan; Regular: 25% * premi
    adb_f = 'IF(C7="Single",C6*0.25/IF(C8="20 Tahun",20,30),C6*0.25)'
    sb.add_row(20, [(1, 'MANFAAT ADB & TPD /TAHUN:', 10), (2, adb_f, 9, 'f'), (3, '25% Premi Dasar Tahunan', 11)])

    # Row 21: Total ADB
    total_adb_f = 'C20*IF(C8="20 Tahun",20,30)'
    sb.add_row(21, [(1, 'TOTAL ADB & TPD:', 10), (2, total_adb_f, 9, 'f')])

    # Row 22: blank
    sb.add_row(22, [])

    # Row 23: TOTAL SELURUH MANFAAT
    total_all_f = 'C19+C21'
    sb.add_row(23, [(1, 'TOTAL SELURUH MANFAAT:', 7), (2, total_all_f, 9, 'f')])

    # Row 25: Notes
    sb.add_row(25, [(1, 'CATATAN:', 7)])
    sb.add_row(26, [(1, '- Pilih dari dropdown di sel KUNING untuk melihat hasil otomatis', 11)])
    sb.add_row(27, [(1, '- Gender tidak mempengaruhi benefit MDWA (unisex rate)', 11)])
    sb.add_row(28, [(1, '- Plan A: Tahapan dibayar tahunan dari tahun mulai s/d akhir coverage', 11)])
    sb.add_row(29, [(1, '- Banding 2: Premi >= 500 Juta (Single) atau >= 100 Juta (Regular)', 11)])

    return sb.build()


# ==============================================================
# PLAN B SHEET - BEASISWA
# ==============================================================
def build_plan_b_sheet(ss_map):
    """PLAN B sheet with own inputs and lookup table for maturity percentages.
    
    Input cells (YELLOW):
    - C4: Jenis Kelamin (Pria/Wanita)
    - C5: Usia Masuk
    - C6: Premi Tahunan (IDR)
    - C7: Masa Bayar (Single/5 Tahun/10 Tahun)
    - C8: Masa Pertanggungan (5-15,20,30 Tahun) - varies by payment term
    
    Lookup table rows 50-80 for maturity percentages.
    """
    sb = SheetBuilder(ss_map)
    sb.set_col_width(1, 8)
    sb.set_col_width(2, 32)
    sb.set_col_width(3, 28)
    sb.set_col_width(4, 30)
    sb.set_col_width(5, 20)

    # Title
    sb.add_row(1, [(1, 'PLAN B - BEASISWA', 3)], height=28)
    sb.add_merge('B1:D1')
    sb.add_row(2, [(1, 'Manfaat Lump Sum di Akhir Masa Pertanggungan (Maturity)', 11)])

    # INPUT section
    sb.add_row(3, [(1, 'INPUT DATA', 5), (2, '', 5), (3, '', 5)])
    sb.add_merge('B3:D3')

    sb.add_row(4, [(1, 'JENIS KELAMIN:', 10), (2, 'Pria', 8)])
    sb.add_data_validation('C4', '"Pria,Wanita"')

    sb.add_row(5, [(1, 'USIA MASUK:', 10), (2, 30, 8, 'n')])

    sb.add_row(6, [(1, 'PREMI TAHUNAN (IDR):', 10), (2, 100000000, 8, 'n')])

    sb.add_row(7, [(1, 'MASA BAYAR:', 10), (2, 'Single', 8)])
    sb.add_data_validation('C7', '"Single,5 Tahun,10 Tahun"')

    sb.add_row(8, [(1, 'MASA PERTANGGUNGAN:', 10), (2, '10 Tahun', 8)])
    sb.add_data_validation('C8', '"5 Tahun,6 Tahun,7 Tahun,8 Tahun,9 Tahun,10 Tahun,11 Tahun,12 Tahun,13 Tahun,14 Tahun,15 Tahun,20 Tahun,30 Tahun"')

    sb.add_row(9, [])

    # HASIL
    sb.add_row(10, [(1, 'HASIL PERHITUNGAN', 5), (2, '', 5), (3, '', 5)])
    sb.add_merge('B10:D10')

    # Eligibility
    elig_f = 'IF(C7="Single",IF(C5<=85,"ELIGIBLE","TIDAK ELIGIBLE - Usia maks 85"),IF(C5<=70,"ELIGIBLE","TIDAK ELIGIBLE - Usia maks 70"))'
    sb.add_row(11, [(1, 'STATUS KELAYAKAN:', 10), (2, elig_f, 12, 'f')])

    # Banding
    banding_f = 'IF(C7="Single",IF(C6>=500000000,"Banding 2","Banding 1"),IF(C6>=100000000,"Banding 2","Banding 1"))'
    sb.add_row(12, [(1, 'LEVEL BANDING:', 10), (2, banding_f, 12, 'f')])

    # Maturity % - uses INDEX/MATCH on lookup table rows 50+
    # Lookup key in A50:A73 = "PayTerm_Coverage" e.g. "Single_5", "5_10", "10_15"
    # Column B50:B73 = maturity % (Banding 1)
    # Column C50:C73 = extra % (Banding 2)
    # Build the key: e.g. IF(C7="Single","Single",IF(C7="5 Tahun","5","10"))&"_"&LEFT(C8,FIND(" ",C8)-1)
    # Simpler: use nested IF since we know all values
    
    # Maturity % Banding 1 - giant nested IF
    # Single: {5:110,6:115,7:120,8:125,9:130,10:135,11:140,12:145,13:150,14:155,15:160,20:200,30:325}
    # 5yr: {10:600,11:625,12:650,13:675,14:700,15:725,20:950,30:1625}
    # 10yr: {15:1350,20:1800,30:3250}
    # We use INDEX/MATCH with the lookup table
    mat_pct_f = 'INDEX(B51:B74,MATCH(IF(C7="Single","Single",IF(C7="5 Tahun","5","10"))&"_"&LEFT(C8,FIND(" ",C8)-1),A51:A74,0))'
    sb.add_row(13, [(1, 'MANFAAT BEASISWA (%):', 10), (2, mat_pct_f, 9, 'f'), (3, '% dari Premi (Banding 1)', 11)])

    # Extra %
    extra_pct_f = 'IF(C12="Banding 2",INDEX(C51:C74,MATCH(IF(C7="Single","Single",IF(C7="5 Tahun","5","10"))&"_"&LEFT(C8,FIND(" ",C8)-1),A51:A74,0)),0)'
    sb.add_row(14, [(1, 'BANDING 2 EXTRA (%):', 10), (2, extra_pct_f, 9, 'f'), (3, 'Tambahan jika Banding 2', 11)])

    # Total %
    total_pct_f = 'C13+C14'
    sb.add_row(15, [(1, 'TOTAL MANFAAT BEASISWA (%):', 10), (2, total_pct_f, 9, 'f')])

    # Nominal lump sum
    nom_f = 'C15*C6/100'
    sb.add_row(16, [(1, 'MANFAAT BEASISWA (LUMP SUM):', 10), (2, nom_f, 9, 'f'), (3, 'Dibayar di akhir coverage', 11)])

    # ADB & TPD
    adb_f = 'IF(C7="Single",C6*0.25/VALUE(LEFT(C8,FIND(" ",C8)-1)),C6*0.25)'
    sb.add_row(17, [(1, 'MANFAAT ADB & TPD /TAHUN:', 10), (2, adb_f, 9, 'f'), (3, '25% Premi Dasar Tahunan', 11)])

    # Total ADB
    total_adb_f = 'C17*VALUE(LEFT(C8,FIND(" ",C8)-1))'
    sb.add_row(18, [(1, 'TOTAL ADB & TPD:', 10), (2, total_adb_f, 9, 'f')])

    sb.add_row(19, [])

    # Total benefit
    total_f = 'C16+C18'
    sb.add_row(20, [(1, 'TOTAL MANFAAT:', 7), (2, total_f, 9, 'f')])

    # Notes
    sb.add_row(22, [(1, 'CATATAN:', 7)])
    sb.add_row(23, [(1, '- Pilih dari dropdown di sel KUNING untuk melihat hasil otomatis', 11)])
    sb.add_row(24, [(1, '- Plan B: Manfaat lump sum dibayar di akhir masa pertanggungan', 11)])
    sb.add_row(25, [(1, '- Single: coverage 5-15/20/30 thn; 5 Tahun: 10-15/20/30 thn; 10 Tahun: 15/20/30 thn', 11)])
    sb.add_row(26, [(1, '- Banding 2: Premi >= 500 Juta (Single) atau >= 100 Juta (Regular)', 11)])

    # ============ LOOKUP TABLE (hidden rows 50+) ============
    sb.add_row(49, [(0, 'TABEL REFERENSI (JANGAN DIUBAH)', 7)])
    sb.add_row(50, [(0, 'Key', 1), (1, 'Maturity % (B1)', 1), (2, 'Extra % (B2)', 1)])

    # Plan B lookup data - starts at row 51
    plan_b_lookup = [
        # Single Premium coverage options
        ('Single_5', 110, 1.5),
        ('Single_6', 115, 1.8),
        ('Single_7', 120, 2.1),
        ('Single_8', 125, 2.4),
        ('Single_9', 130, 2.7),
        ('Single_10', 135, 3),
        ('Single_11', 140, 3.3),
        ('Single_12', 145, 3.6),
        ('Single_13', 150, 3.9),
        ('Single_14', 155, 4.2),
        ('Single_15', 160, 4.5),
        ('Single_20', 200, 6),
        ('Single_30', 325, 9),
        # PPP 5 Tahun
        ('5_10', 600, 15),
        ('5_11', 625, 16.5),
        ('5_12', 650, 18),
        ('5_13', 675, 19.5),
        ('5_14', 700, 21),
        ('5_15', 725, 22.5),
        ('5_20', 950, 30),
        ('5_30', 1625, 45),
        # PPP 10 Tahun
        ('10_15', 1350, 45),
        ('10_20', 1800, 60),
        ('10_30', 3250, 90),
    ]

    for i, (key, mat, extra) in enumerate(plan_b_lookup):
        row = 51 + i
        sb.add_row(row, [(0, key, 2), (1, mat, 6, 'n'), (2, extra, 6, 'n')])

    return sb.build()


# ==============================================================
# PLAN C SHEET - COMBO
# ==============================================================
def build_plan_c_sheet(ss_map):
    """PLAN C sheet with own inputs and lookup tables.
    
    Input cells (YELLOW):
    - C4: Jenis Kelamin
    - C5: Usia Masuk
    - C6: Premi Tahunan (IDR)
    - C7: Masa Bayar (Single/5 Tahun/10 Tahun)
    - C8: Masa Pertanggungan (10/15/20/30 Tahun)
    
    Results: Tahapan + Maturity + ADB
    Lookup table rows 50+ for percentages.
    """
    sb = SheetBuilder(ss_map)
    sb.set_col_width(1, 8)
    sb.set_col_width(2, 36)
    sb.set_col_width(3, 28)
    sb.set_col_width(4, 30)
    sb.set_col_width(5, 20)

    # Title
    sb.add_row(1, [(1, 'PLAN C - COMBO', 3)], height=28)
    sb.add_merge('B1:D1')
    sb.add_row(2, [(1, 'Manfaat Tahapan Tahunan + Lump Sum Jatuh Tempo', 11)])

    # INPUT section
    sb.add_row(3, [(1, 'INPUT DATA', 5), (2, '', 5), (3, '', 5)])
    sb.add_merge('B3:D3')

    sb.add_row(4, [(1, 'JENIS KELAMIN:', 10), (2, 'Pria', 8)])
    sb.add_data_validation('C4', '"Pria,Wanita"')

    sb.add_row(5, [(1, 'USIA MASUK:', 10), (2, 30, 8, 'n')])

    sb.add_row(6, [(1, 'PREMI TAHUNAN (IDR):', 10), (2, 100000000, 8, 'n')])

    sb.add_row(7, [(1, 'MASA BAYAR:', 10), (2, 'Single', 8)])
    sb.add_data_validation('C7', '"Single,5 Tahun,10 Tahun"')

    sb.add_row(8, [(1, 'MASA PERTANGGUNGAN:', 10), (2, '20 Tahun', 8)])
    sb.add_data_validation('C8', '"10 Tahun,15 Tahun,20 Tahun,30 Tahun"')

    sb.add_row(9, [])

    # HASIL
    sb.add_row(10, [(1, 'HASIL PERHITUNGAN', 5), (2, '', 5), (3, '', 5)])
    sb.add_merge('B10:D10')

    # Eligibility
    elig_f = 'IF(C7="Single",IF(C5<=85,"ELIGIBLE","TIDAK ELIGIBLE - Usia maks 85"),IF(C5<=70,"ELIGIBLE","TIDAK ELIGIBLE - Usia maks 70"))'
    sb.add_row(11, [(1, 'STATUS KELAYAKAN:', 10), (2, elig_f, 12, 'f')])

    # Banding
    banding_f = 'IF(C7="Single",IF(C6>=500000000,"Banding 2","Banding 1"),IF(C6>=100000000,"Banding 2","Banding 1"))'
    sb.add_row(12, [(1, 'LEVEL BANDING:', 10), (2, banding_f, 12, 'f')])

    # Tahapan % - INDEX/MATCH from lookup table
    # Key format: "Single_10", "5_15", "10_20" etc.
    tahapan_pct_f = 'INDEX(B51:B61,MATCH(IF(C7="Single","Single",IF(C7="5 Tahun","5","10"))&"_"&LEFT(C8,FIND(" ",C8)-1),A51:A61,0))'
    sb.add_row(13, [(1, 'MANFAAT TAHAPAN TAHUNAN (%):', 10), (2, tahapan_pct_f, 9, 'f'), (3, '% dari Premi (Banding 1)', 11)])

    # Tahapan nominal
    tahapan_nom_f = 'C13*C6/100'
    sb.add_row(14, [(1, 'MANFAAT TAHAPAN TAHUNAN (IDR):', 10), (2, tahapan_nom_f, 9, 'f'), (3, 'Per tahun', 11)])

    # Maturity %
    mat_pct_f = 'INDEX(C51:C61,MATCH(IF(C7="Single","Single",IF(C7="5 Tahun","5","10"))&"_"&LEFT(C8,FIND(" ",C8)-1),A51:A61,0))'
    sb.add_row(15, [(1, 'MANFAAT AKHIR MASA PERTANGGUNGAN (%):', 10), (2, mat_pct_f, 9, 'f'), (3, '% dari Premi (Banding 1)', 11)])

    # Maturity nominal
    mat_nom_f = 'C15*C6/100'
    sb.add_row(16, [(1, 'MANFAAT AKHIR MASA PERTANGGUNGAN (IDR):', 10), (2, mat_nom_f, 9, 'f')])

    # Banding 2 Extra (maturity only)
    extra_pct_f = 'IF(C12="Banding 2",INDEX(D51:D61,MATCH(IF(C7="Single","Single",IF(C7="5 Tahun","5","10"))&"_"&LEFT(C8,FIND(" ",C8)-1),A51:A61,0)),0)'
    sb.add_row(17, [(1, 'BANDING 2 EXTRA MATURITY (%):', 10), (2, extra_pct_f, 9, 'f'), (3, 'Tambahan jika Banding 2', 11)])

    # Extra nominal
    extra_nom_f = 'C17*C6/100'
    sb.add_row(18, [(1, 'BANDING 2 EXTRA (IDR):', 10), (2, extra_nom_f, 9, 'f')])

    # Mulai Pembayaran Tahapan
    # Single: yr 1, 5yr: yr 5, 10yr: yr 10
    mulai_f = 'IF(C7="Single","Tahun ke-1",IF(C7="5 Tahun","Tahun ke-5","Tahun ke-10"))'
    sb.add_row(19, [(1, 'MULAI PEMBAYARAN TAHAPAN:', 10), (2, mulai_f, 12, 'f')])

    # Jumlah tahun pembayaran tahapan
    # coverage - start + 1: Single: cov-0=cov, 5yr: cov-4, 10yr: cov-9
    jml_f = 'IF(C7="Single",VALUE(LEFT(C8,FIND(" ",C8)-1)),IF(C7="5 Tahun",VALUE(LEFT(C8,FIND(" ",C8)-1))-4,VALUE(LEFT(C8,FIND(" ",C8)-1))-9))'
    sb.add_row(20, [(1, 'JUMLAH TAHUN PEMBAYARAN:', 10), (2, jml_f, 9, 'f'), (3, 'tahun', 11)])

    # Total Tahapan
    total_tahapan_f = 'C14*C20'
    sb.add_row(21, [(1, 'TOTAL TAHAPAN (SELURUH PERIODE):', 10), (2, total_tahapan_f, 9, 'f')])

    # ADB & TPD
    adb_f = 'IF(C7="Single",C6*0.25/VALUE(LEFT(C8,FIND(" ",C8)-1)),C6*0.25)'
    sb.add_row(22, [(1, 'MANFAAT ADB & TPD /TAHUN:', 10), (2, adb_f, 9, 'f'), (3, '25% Premi Dasar Tahunan', 11)])

    # Total ADB
    total_adb_f = 'C22*VALUE(LEFT(C8,FIND(" ",C8)-1))'
    sb.add_row(23, [(1, 'TOTAL ADB & TPD:', 10), (2, total_adb_f, 9, 'f')])

    sb.add_row(24, [])

    # Total benefit
    total_f = 'C21+C16+C18+C23'
    sb.add_row(25, [(1, 'TOTAL SELURUH MANFAAT:', 7), (2, total_f, 9, 'f')])

    # Notes
    sb.add_row(27, [(1, 'CATATAN:', 7)])
    sb.add_row(28, [(1, '- Pilih dari dropdown di sel KUNING untuk melihat hasil otomatis', 11)])
    sb.add_row(29, [(1, '- Plan C: Gabungan tahapan tahunan + lump sum di akhir', 11)])
    sb.add_row(30, [(1, '- Single: mulai tahun ke-1; 5 Tahun: mulai tahun ke-5; 10 Tahun: mulai tahun ke-10', 11)])
    sb.add_row(31, [(1, '- Banding 2: Premi >= 500 Juta (Single) atau >= 100 Juta (Regular)', 11)])

    # ============ LOOKUP TABLE (rows 50+) ============
    sb.add_row(49, [(0, 'TABEL REFERENSI (JANGAN DIUBAH)', 7)])
    # Columns: A=Key, B=Tahapan%, C=Maturity%, D=Extra Maturity% (B2)
    sb.add_row(50, [(0, 'Key', 1), (1, 'Tahapan %', 1), (2, 'Maturity %', 1), (3, 'Extra Mat %', 1)])

    plan_c_lookup = [
        # Single
        ('Single_10', 2, 105, 3),
        ('Single_15', 2, 120, 5),
        ('Single_20', 2, 135, 6),
        ('Single_30', 2, 170, 9),
        # PPP 5
        ('5_10', 13.5, 510, 15),
        ('5_15', 17.5, 510, 23),
        ('5_20', 20, 510, 30),
        ('5_30', 20, 510, 45),
        # PPP 10
        ('10_15', 40, 1000, 45),
        ('10_20', 40, 1100, 60),
        ('10_30', 40, 1200, 90),
    ]

    for i, (key, tahapan, maturity, extra) in enumerate(plan_c_lookup):
        row = 51 + i
        sb.add_row(row, [(0, key, 2), (1, tahapan, 6, 'n'), (2, maturity, 6, 'n'), (3, extra, 6, 'n')])

    return sb.build()


# ==============================================================
# KOMISI SHEET
# ==============================================================
def build_komisi_sheet(ss_map):
    """Commission structure for all 3 plans."""
    sb = SheetBuilder(ss_map)
    sb.set_col_width(1, 8)
    sb.set_col_width(2, 20)
    sb.set_col_width(3, 18)
    sb.set_col_width(4, 15)
    sb.set_col_width(5, 15)

    r = 1
    sb.add_row(r, [(1, 'STRUKTUR KOMISI MDWA', 5), (2, '', 5), (3, '', 5), (4, '', 5)])
    sb.add_merge('B1:E1')
    r += 2

    # Plan A
    sb.add_row(r, [(1, 'PLAN A - ANUITAS', 7)])
    r += 1
    sb.add_row(r, [(1, 'Masa Bayar', 1), (2, 'Coverage', 1), (3, 'Tahun 1', 1), (4, 'Tahun 2', 1)])
    r += 1
    for term, cov, y1, y2 in [
        ('Single', '20 thn', '4%', '-'),
        ('Single', '30 thn', '4%', '-'),
        ('5 Tahun', '20 thn', '5%', '5%'),
        ('5 Tahun', '30 thn', '5%', '5%'),
        ('10 Tahun', '20 thn', '15%', '10%'),
        ('10 Tahun', '30 thn', '15%', '10%'),
    ]:
        sb.add_row(r, [(1, term, 2), (2, cov, 2), (3, y1, 2), (4, y2, 2)])
        r += 1
    r += 1

    # Plan B
    sb.add_row(r, [(1, 'PLAN B - BEASISWA', 7)])
    r += 1
    sb.add_row(r, [(1, 'Masa Bayar', 1), (2, 'Coverage', 1), (3, 'Tahun 1', 1), (4, 'Tahun 2', 1)])
    r += 1
    for term, cov, y1, y2 in [
        ('Single', '5-10 thn', '2%', '-'),
        ('Single', '11-15 thn', '5%', '-'),
        ('Single', '20 thn', '5%', '-'),
        ('Single', '30 thn', '6%', '-'),
        ('5 Tahun', '5-15 thn', '7.5%', '5%'),
        ('5 Tahun', '20 thn', '10%', '5%'),
        ('5 Tahun', '30 thn', '10%', '5%'),
        ('10 Tahun', 'semua', '15%', '10%'),
    ]:
        sb.add_row(r, [(1, term, 2), (2, cov, 2), (3, y1, 2), (4, y2, 2)])
        r += 1
    r += 1

    # Plan C
    sb.add_row(r, [(1, 'PLAN C - COMBO', 7)])
    r += 1
    sb.add_row(r, [(1, 'Masa Bayar', 1), (2, 'Coverage', 1), (3, 'Tahun 1', 1), (4, 'Tahun 2', 1)])
    r += 1
    for term, cov, y1, y2 in [
        ('Single', '10 thn', '2%', '-'),
        ('Single', '15 thn', '2%', '-'),
        ('Single', '20 thn', '4%', '-'),
        ('Single', '30 thn', '6%', '-'),
        ('5 Tahun', '10 thn', '5%', '5%'),
        ('5 Tahun', '15 thn', '7.5%', '5%'),
        ('5 Tahun', '20 thn', '10%', '5%'),
        ('5 Tahun', '30 thn', '12.5%', '5%'),
        ('10 Tahun', 'semua', '15%', '10%'),
    ]:
        sb.add_row(r, [(1, term, 2), (2, cov, 2), (3, y1, 2), (4, y2, 2)])
        r += 1

    return sb.build()


# ==============================================================
# PERBANDINGAN SHEET
# ==============================================================
def build_perbandingan_sheet(ss_map):
    """Comparison vs competitors."""
    sb = SheetBuilder(ss_map)
    sb.set_col_width(1, 8)
    for i in range(2, 9):
        sb.set_col_width(i, 18)

    r = 1
    sb.add_row(r, [(1, 'PERBANDINGAN vs KOMPETITOR', 5)])
    sb.add_merge('B1:H1')
    r += 2

    sb.add_row(r, [(1, 'Contoh: PPP 5 Tahun, Banding 2, Premi Tahunan IDR 100 Juta', 7)])
    r += 2

    sb.add_row(r, [(1, 'Fitur', 1), (2, 'MDWA 15yr', 1), (3, 'MDWA 20yr', 1),
                   (4, 'MSP 15yr', 1), (5, 'MSP 20yr', 1), (6, 'FI 15yr', 1), (7, 'FI 20yr', 1)])
    r += 1
    sb.add_row(r, [(1, 'Tahapan/Tahun', 2), (2, '17.5M x 11', 2), (3, '20M x 16', 2),
                   (4, '17.5M x 11', 2), (5, '20M x 16', 2), (6, '20M x 10', 2), (7, '20M x 15', 2)])
    r += 1
    sb.add_row(r, [(1, 'Total Tahapan', 2), (2, '192.5M', 2), (3, '320M', 2),
                   (4, '192.5M', 2), (5, '320M', 2), (6, '200M', 2), (7, '300M', 2)])
    r += 1
    sb.add_row(r, [(1, 'Maturity', 2), (2, '532.5M', 2), (3, '540M', 2),
                   (4, '500M', 2), (5, '500M', 2), (6, '500M', 2), (7, '525M', 2)])
    r += 1
    sb.add_row(r, [(1, 'Total Benefit', 4), (2, '725M', 4), (3, '860M', 4),
                   (4, '692.5M', 4), (5, '820M', 4), (6, '700M', 4), (7, '825M', 4)])
    r += 1
    sb.add_row(r, [(1, 'Rasio Benefit', 4), (2, '145%', 4), (3, '172%', 4),
                   (4, '139%', 4), (5, '164%', 4), (6, '140%', 4), (7, '165%', 4)])
    r += 2

    sb.add_row(r, [(1, 'Catatan:', 7)])
    r += 1
    sb.add_row(r, [(1, 'MSP = Manulife Saver Plus (produk lama)', 0)])
    r += 1
    sb.add_row(r, [(1, 'FI = Flexi Income (kompetitor)', 0)])
    r += 1
    sb.add_row(r, [(1, 'Rasio = Total Benefit / Total Premi (5 thn x 100 Juta = 500 Juta)', 0)])
    r += 1
    sb.add_row(r, [(1, 'MDWA memberikan rasio benefit tertinggi.', 7)])

    return sb.build()


# ==============================================================
# RINGKASAN SHEET
# ==============================================================
def build_ringkasan_sheet(ss_map):
    """Summary/overview sheet comparing all 3 plans side by side."""
    sb = SheetBuilder(ss_map)
    sb.set_col_width(1, 8)
    sb.set_col_width(2, 28)
    sb.set_col_width(3, 32)
    sb.set_col_width(4, 32)
    sb.set_col_width(5, 32)

    r = 1
    sb.add_row(r, [(1, 'RINGKASAN - MANULIFE DYNAMIC WEALTH ASSURANCE', 5)])
    sb.add_merge('B1:E1')
    r += 2

    # Plan comparison header
    sb.add_row(r, [(1, 'Fitur', 1), (2, 'Plan A - Anuitas', 1), (3, 'Plan B - Beasiswa', 1), (4, 'Plan C - Combo', 1)])
    r += 1
    rows_data = [
        ('Konsep', 'Arus kas tahunan stabil (pensiun)', 'Lump sum di akhir (pendidikan)', 'Tahapan + lump sum (fleksibel)'),
        ('Manfaat Utama', 'Tahapan tahunan (annual payout)', 'Lump sum saat jatuh tempo', 'Tahapan + jatuh tempo'),
        ('Masa Pertanggungan', '20 / 30 tahun', '5-15 / 20 / 30 tahun', '10 / 15 / 20 / 30 tahun'),
        ('Masa Bayar Premi', 'Single / 5 / 10 tahun', 'Single / 5 / 10 tahun', 'Single / 5 / 10 tahun'),
        ('Usia Masuk (Single)', '30 hari - 85 tahun', '30 hari - 85 tahun', '30 hari - 85 tahun'),
        ('Usia Masuk (Regular)', '30 hari - 70 tahun', '30 hari - 70 tahun', '30 hari - 70 tahun'),
        ('Min Premi Single', 'IDR 50 Juta / USD 5,000', 'IDR 50 Juta / USD 5,000', 'IDR 50 Juta / USD 5,000'),
        ('Min Premi Regular', 'IDR 24 Juta / USD 2,400', 'IDR 24 Juta / USD 2,400', 'IDR 24 Juta / USD 2,400'),
        ('ADB & TPD', '25% Premi Dasar Tahunan', '25% Premi Dasar Tahunan', '25% Premi Dasar Tahunan'),
        ('Banding 2 Trigger', 'Single>=500M / Reg>=100M', 'Single>=500M / Reg>=100M', 'Single>=500M / Reg>=100M'),
    ]
    for label, a, b, c in rows_data:
        sb.add_row(r, [(1, label, 4), (2, a, 2), (3, b, 2), (4, c, 2)])
        r += 1

    r += 1
    sb.add_row(r, [(1, 'SPESIFIKASI UMUM', 5)])
    sb.add_merge('B' + str(r) + ':E' + str(r))
    r += 1
    specs = [
        ('Mode Pembayaran', 'Tahunan (100%), Semesteran (52.5%/51.25%), Triwulanan (27.5%/26.5%), Bulanan (9.5%/9%)'),
        ('Underwriting', 'Guaranteed Issuance Offering'),
        ('Riders', 'Advanced Life Protector Plus, Manulife Payor Benefit Plus, Manulife Waiver of Premium Plus'),
        ('Mata Uang', 'IDR dan USD'),
    ]
    for label, val in specs:
        sb.add_row(r, [(1, label, 4), (2, val, 2)])
        sb.add_merge('C' + str(r) + ':E' + str(r))
        r += 1

    r += 1
    sb.add_row(r, [(1, 'BANDING PREMIUM', 5)])
    sb.add_merge('B' + str(r) + ':E' + str(r))
    r += 1
    sb.add_row(r, [(1, 'Banding', 1), (2, 'Single Premium', 1), (3, 'Regular Premium', 1)])
    r += 1
    sb.add_row(r, [(1, 'Banding 1 (IDR)', 2), (2, 'IDR 50M - <500M', 2), (3, 'IDR 24M - <100M', 2)])
    r += 1
    sb.add_row(r, [(1, 'Banding 2 (IDR)', 2), (2, '>= IDR 500M', 2), (3, '>= IDR 100M', 2)])
    r += 1
    sb.add_row(r, [(1, 'Banding 1 (USD)', 2), (2, 'USD 5,000 - <50,000', 2), (3, 'USD 2,400 - <10,000', 2)])
    r += 1
    sb.add_row(r, [(1, 'Banding 2 (USD)', 2), (2, '>= USD 50,000', 2), (3, '>= USD 10,000', 2)])

    r += 2
    sb.add_row(r, [(1, 'Gunakan sheet PLAN A, PLAN B, PLAN C untuk kalkulasi manfaat.', 7)])
    r += 1
    sb.add_row(r, [(1, 'Pilih Gender dan Usia di sel KUNING pada masing-masing sheet Plan.', 11)])

    return sb.build()


# ==============================================================
# MAIN - Build and package XLSX
# ==============================================================
def main():
    print('Generating MDWA_Kalkulator_v2.xlsx (FIXED dropdown + yellow cells)...')

    # Collect shared strings
    all_strings = []
    seen = set()

    def add_str(s):
        if s not in seen:
            seen.add(s)
            all_strings.append(s)

    # All unique strings across sheets
    strings_list = [
        # Plan A
        'PLAN A - ANUITAS', 'Manfaat Tahapan Tahunan (Annual Payout)',
        'INPUT DATA', 'JENIS KELAMIN:', 'Pria', 'Wanita',
        'USIA MASUK:', 'PREMI TAHUNAN (IDR):', 'MASA BAYAR:',
        'Single', '5 Tahun', '10 Tahun',
        'MASA PERTANGGUNGAN:', '20 Tahun', '30 Tahun',
        'HASIL PERHITUNGAN', 'STATUS KELAYAKAN:', 'LEVEL BANDING:',
        'MANFAAT TAHAPAN TAHUNAN (%):', '% dari Premi (Banding 1)',
        'BANDING 2 EXTRA (%):', 'Tambahan jika Banding 2',
        'TOTAL TAHAPAN % PER TAHUN:', 'MANFAAT TAHAPAN TAHUNAN (IDR):',
        'Per tahun', 'MULAI PEMBAYARAN:', 'JUMLAH TAHUN PEMBAYARAN:',
        'tahun', 'TOTAL TAHAPAN (SELURUH PERIODE):',
        'MANFAAT ADB & TPD /TAHUN:', '25% Premi Dasar Tahunan',
        'TOTAL ADB & TPD:', 'TOTAL SELURUH MANFAAT:',
        'CATATAN:',
        '- Pilih dari dropdown di sel KUNING untuk melihat hasil otomatis',
        '- Gender tidak mempengaruhi benefit MDWA (unisex rate)',
        '- Plan A: Tahapan dibayar tahunan dari tahun mulai s/d akhir coverage',
        '- Banding 2: Premi >= 500 Juta (Single) atau >= 100 Juta (Regular)',
        # Plan B
        'PLAN B - BEASISWA', 'Manfaat Lump Sum di Akhir Masa Pertanggungan (Maturity)',
        '5 Tahun', '6 Tahun', '7 Tahun', '8 Tahun', '9 Tahun',
        '10 Tahun', '11 Tahun', '12 Tahun', '13 Tahun', '14 Tahun', '15 Tahun',
        'MANFAAT BEASISWA (%):', 'TOTAL MANFAAT BEASISWA (%):',
        'MANFAAT BEASISWA (LUMP SUM):', 'Dibayar di akhir coverage',
        'TOTAL MANFAAT:',
        '- Plan B: Manfaat lump sum dibayar di akhir masa pertanggungan',
        '- Single: coverage 5-15/20/30 thn; 5 Tahun: 10-15/20/30 thn; 10 Tahun: 15/20/30 thn',
        'TABEL REFERENSI (JANGAN DIUBAH)', 'Key', 'Maturity % (B1)', 'Extra % (B2)',
        'Single_5', 'Single_6', 'Single_7', 'Single_8', 'Single_9',
        'Single_10', 'Single_11', 'Single_12', 'Single_13', 'Single_14', 'Single_15',
        'Single_20', 'Single_30', '5_10', '5_11', '5_12', '5_13', '5_14', '5_15',
        '5_20', '5_30', '10_15', '10_20', '10_30',
        # Plan C
        'PLAN C - COMBO', 'Manfaat Tahapan Tahunan + Lump Sum Jatuh Tempo',
        'MANFAAT AKHIR MASA PERTANGGUNGAN (%):', 'MANFAAT AKHIR MASA PERTANGGUNGAN (IDR):',
        'BANDING 2 EXTRA MATURITY (%):', 'BANDING 2 EXTRA (IDR):',
        'MULAI PEMBAYARAN TAHAPAN:',
        '- Plan C: Gabungan tahapan tahunan + lump sum di akhir',
        '- Single: mulai tahun ke-1; 5 Tahun: mulai tahun ke-5; 10 Tahun: mulai tahun ke-10',
        'Tahapan %', 'Maturity %', 'Extra Mat %',
        # Komisi
        'STRUKTUR KOMISI MDWA', 'PLAN A - ANUITAS', 'Masa Bayar', 'Coverage',
        'Tahun 1', 'Tahun 2',
        '20 thn', '30 thn', '4%', '5%', '15%', '10%', '-',
        'PLAN B - BEASISWA', '5-10 thn', '11-15 thn', '2%', '6%', '7.5%',
        'PLAN C - COMBO', '10 thn', '15 thn', '12.5%', 'semua',
        # Perbandingan
        'PERBANDINGAN vs KOMPETITOR',
        'Contoh: PPP 5 Tahun, Banding 2, Premi Tahunan IDR 100 Juta',
        'Fitur', 'MDWA 15yr', 'MDWA 20yr', 'MSP 15yr', 'MSP 20yr', 'FI 15yr', 'FI 20yr',
        'Tahapan/Tahun', '17.5M x 11', '20M x 16', '20M x 10', '20M x 15',
        'Total Tahapan', '192.5M', '320M', '200M', '300M',
        'Maturity', '532.5M', '540M', '500M', '525M',
        'Total Benefit', '725M', '860M', '692.5M', '820M', '700M', '825M',
        'Rasio Benefit', '145%', '172%', '139%', '164%', '140%', '165%',
        'Catatan:',
        'MSP = Manulife Saver Plus (produk lama)',
        'FI = Flexi Income (kompetitor)',
        'Rasio = Total Benefit / Total Premi (5 thn x 100 Juta = 500 Juta)',
        'MDWA memberikan rasio benefit tertinggi.',
        # Ringkasan
        'RINGKASAN - MANULIFE DYNAMIC WEALTH ASSURANCE',
        'Plan A - Anuitas', 'Plan B - Beasiswa', 'Plan C - Combo',
        'Konsep', 'Arus kas tahunan stabil (pensiun)',
        'Lump sum di akhir (pendidikan)', 'Tahapan + lump sum (fleksibel)',
        'Manfaat Utama', 'Tahapan tahunan (annual payout)',
        'Lump sum saat jatuh tempo', 'Tahapan + jatuh tempo',
        'Masa Pertanggungan', '20 / 30 tahun', '5-15 / 20 / 30 tahun', '10 / 15 / 20 / 30 tahun',
        'Masa Bayar Premi', 'Single / 5 / 10 tahun',
        'Usia Masuk (Single)', '30 hari - 85 tahun',
        'Usia Masuk (Regular)', '30 hari - 70 tahun',
        'Min Premi Single', 'IDR 50 Juta / USD 5,000',
        'Min Premi Regular', 'IDR 24 Juta / USD 2,400',
        'ADB & TPD', '25% Premi Dasar Tahunan',
        'Banding 2 Trigger', 'Single>=500M / Reg>=100M',
        'SPESIFIKASI UMUM',
        'Mode Pembayaran', 'Tahunan (100%), Semesteran (52.5%/51.25%), Triwulanan (27.5%/26.5%), Bulanan (9.5%/9%)',
        'Underwriting', 'Guaranteed Issuance Offering',
        'Riders', 'Advanced Life Protector Plus, Manulife Payor Benefit Plus, Manulife Waiver of Premium Plus',
        'Mata Uang', 'IDR dan USD',
        'BANDING PREMIUM', 'Banding', 'Single Premium', 'Regular Premium',
        'Banding 1 (IDR)', 'IDR 50M - <500M', 'IDR 24M - <100M',
        'Banding 2 (IDR)', '>= IDR 500M', '>= IDR 100M',
        'Banding 1 (USD)', 'USD 5,000 - <50,000', 'USD 2,400 - <10,000',
        'Banding 2 (USD)', '>= USD 50,000', '>= USD 10,000',
        'Gunakan sheet PLAN A, PLAN B, PLAN C untuk kalkulasi manfaat.',
        'Pilih Gender dan Usia di sel KUNING pada masing-masing sheet Plan.',
    ]
    for s in strings_list:
        add_str(s)

    ss_map = {s: i for i, s in enumerate(all_strings)}

    print('  Building sheets...')
    sheet1_xml = build_plan_a_sheet(ss_map)
    sheet2_xml = build_plan_b_sheet(ss_map)
    sheet3_xml = build_plan_c_sheet(ss_map)
    sheet4_xml = build_komisi_sheet(ss_map)
    sheet5_xml = build_perbandingan_sheet(ss_map)
    sheet6_xml = build_ringkasan_sheet(ss_map)

    shared_strings_xml = build_shared_strings(all_strings)
    styles_xml = build_styles_xml()
    theme_xml = build_theme_xml()

    sheet_names = ['PLAN A', 'PLAN B', 'PLAN C', 'KOMISI', 'PERBANDINGAN', 'RINGKASAN']
    num_sheets = len(sheet_names)

    # Content Types
    ct = chr(60) + '?xml version="1.0" encoding="UTF-8" standalone="yes"?' + chr(62)
    ct += '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
    ct += '<Default ContentType="application/xml" Extension="xml"/>'
    ct += '<Default ContentType="application/vnd.openxmlformats-package.relationships+xml" Extension="rels"/>'
    for i in range(1, num_sheets + 1):
        ct += '<Override ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml" PartName="/xl/worksheets/sheet' + str(i) + '.xml"/>'
    ct += '<Override ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml" PartName="/xl/sharedStrings.xml"/>'
    ct += '<Override ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml" PartName="/xl/styles.xml"/>'
    ct += '<Override ContentType="application/vnd.openxmlformats-officedocument.theme+xml" PartName="/xl/theme/theme1.xml"/>'
    ct += '<Override ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml" PartName="/xl/workbook.xml"/>'
    ct += '</Types>'

    # Top-level rels
    rels = chr(60) + '?xml version="1.0" encoding="UTF-8" standalone="yes"?' + chr(62)
    rels += '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
    rels += '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>'
    rels += '</Relationships>'

    # Workbook
    wb = chr(60) + '?xml version="1.0" encoding="UTF-8" standalone="yes"?' + chr(62)
    wb += '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"'
    wb += ' xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
    wb += '<workbookPr/><sheets>'
    for i, name in enumerate(sheet_names, 1):
        wb += '<sheet state="visible" name="' + name + '" sheetId="' + str(i) + '" r:id="rId' + str(i + 3) + '"/>'
    wb += '</sheets><definedNames/><calcPr/></workbook>'

    # Workbook rels
    wb_rels = chr(60) + '?xml version="1.0" encoding="UTF-8" standalone="yes"?' + chr(62)
    wb_rels += '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
    wb_rels += '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml"/>'
    wb_rels += '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>'
    wb_rels += '<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/>'
    for i in range(1, num_sheets + 1):
        wb_rels += '<Relationship Id="rId' + str(i + 3) + '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet' + str(i) + '.xml"/>'
    wb_rels += '</Relationships>'

    print('  Packaging xlsx...')
    sheets_xml = [sheet1_xml, sheet2_xml, sheet3_xml, sheet4_xml, sheet5_xml, sheet6_xml]

    with zipfile.ZipFile(OUTPUT_XLSX, 'w', zipfile.ZIP_DEFLATED) as zf:
        zf.writestr('[Content_Types].xml', ct)
        zf.writestr('_rels/.rels', rels)
        zf.writestr('xl/workbook.xml', wb)
        zf.writestr('xl/_rels/workbook.xml.rels', wb_rels)
        zf.writestr('xl/styles.xml', styles_xml)
        zf.writestr('xl/sharedStrings.xml', shared_strings_xml)
        zf.writestr('xl/theme/theme1.xml', theme_xml)
        for i, xml in enumerate(sheets_xml, 1):
            zf.writestr('xl/worksheets/sheet' + str(i) + '.xml', xml)

    print(f'Output: {OUTPUT_XLSX}')

    # Verify
    with zipfile.ZipFile(OUTPUT_XLSX) as z:
        print(f'  Valid xlsx with {len(z.namelist())} files')
        print(f'  Files: {z.namelist()}')

    # Verify formulas and data validation
    with zipfile.ZipFile(OUTPUT_XLSX) as z:
        for i in range(1, 4):
            sheet_data = z.read(f'xl/worksheets/sheet{i}.xml').decode()
            formulas = re.findall(r'<f>(.*?)</f>', sheet_data)
            has_dv = 'dataValidation' in sheet_data
            print(f'  Sheet {i} ({sheet_names[i-1]}): {len(formulas)} formulas, dataValidation={has_dv}')

    # Verify no Chinese chars
    with zipfile.ZipFile(OUTPUT_XLSX) as z:
        content = z.read('xl/sharedStrings.xml').decode()
        has_chinese = any('\u4e00' <= c <= '\u9fff' for c in content)
        print(f'  Chinese characters: {"FOUND - ERROR" if has_chinese else "None - OK"}')

    print('Done!')


if __name__ == "__main__":
    main()
