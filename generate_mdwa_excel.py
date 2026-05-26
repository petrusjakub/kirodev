#!/usr/bin/env python3
"""Generate Kalkulator_MDWA.xlsx - Interactive MDWA calculator with lookup functionality.

Redesigned with KALKULATOR sheet featuring:
- Gender/Age/Premium/Plan/Currency/Payment Term/Coverage Period inputs
- Excel formulas (IF, INDEX/MATCH) for automatic benefit calculation
- Data validation dropdowns for input cells
- Reference data sheets for Plan A, B, C
"""
import zipfile
import os

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
OUTPUT_XLSX = os.path.join(SCRIPT_DIR, 'Kalkulator_MDWA.xlsx')


def xml_escape(s):
    return s.replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;').replace('"', '&quot;')


def col_letter(col_idx):
    result = ''
    idx = col_idx
    while True:
        result = chr(65 + idx % 26) + result
        idx = idx // 26 - 1
        if idx < 0:
            break
    return result


class SheetBuilder:
    """Helper to build worksheet XML with rows, formulas, data validation, merges."""

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
            # Formula cell
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
                s += '<dataValidation type="list" sqref="' + sqref + '"' + blank + ' showInputMessage="1" showErrorMessage="1">'
                s += '<formula1>' + xml_escape(formula1) + '</formula1>'
                s += '</dataValidation>'
            s += '</dataValidations>'
        s += '</worksheet>'
        return s



def build_styles_xml():
    s = chr(60) + '?xml version="1.0" encoding="UTF-8" standalone="yes"?' + chr(62)
    s += '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
    # numFmts
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
    # 6=light yellow (input), 7=light blue (result)
    s += '<fills count="8">'
    s += '<fill><patternFill patternType="none"/></fill>'
    s += '<fill><patternFill patternType="lightGray"/></fill>'
    s += '<fill><patternFill patternType="solid"><fgColor rgb="FF1E7145"/><bgColor rgb="FF1E7145"/></patternFill></fill>'
    s += '<fill><patternFill patternType="solid"><fgColor rgb="FFE8F5EE"/><bgColor rgb="FFE8F5EE"/></patternFill></fill>'
    s += '<fill><patternFill patternType="solid"><fgColor rgb="FFFFFFFF"/><bgColor rgb="FFFFFFFF"/></patternFill></fill>'
    s += '<fill><patternFill patternType="solid"><fgColor rgb="FF00B050"/><bgColor rgb="FF00B050"/></patternFill></fill>'
    s += '<fill><patternFill patternType="solid"><fgColor rgb="FFFFFFCC"/><bgColor rgb="FFFFFFCC"/></patternFill></fill>'
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
    # cellStyleXfs
    s += '<cellStyleXfs count="1">'
    s += '<xf borderId="0" fillId="0" fontId="0" numFmtId="0"/>'
    s += '</cellStyleXfs>'
    # cellXfs  (styles 0-12)
    # 0: default
    # 1: header green bg white bold text centered
    # 2: data cell with border centered
    # 3: title bold green font
    # 4: data cell light green bg
    # 5: section header dark green bg white bold large
    # 6: number format with comma, border, centered
    # 7: bold black left aligned
    # 8: input cell (yellow bg, blue border, bold blue font)
    # 9: result cell (light blue bg, border, number format)
    # 10: label bold black with border
    # 11: note italic gray
    # 12: percentage format with border centered
    s += '<cellXfs count="13">'
    s += '<xf borderId="0" fillId="0" fontId="0" numFmtId="0" xfId="0" applyAlignment="1"><alignment vertical="center" wrapText="1"/></xf>'
    s += '<xf borderId="1" fillId="2" fontId="1" numFmtId="0" xfId="0" applyAlignment="1" applyBorder="1" applyFill="1" applyFont="1"><alignment horizontal="center" vertical="center" wrapText="1"/></xf>'
    s += '<xf borderId="1" fillId="0" fontId="0" numFmtId="0" xfId="0" applyAlignment="1" applyBorder="1"><alignment horizontal="center" vertical="center" wrapText="1"/></xf>'
    s += '<xf borderId="0" fillId="0" fontId="2" numFmtId="0" xfId="0" applyAlignment="1" applyFont="1"><alignment vertical="center" wrapText="1"/></xf>'
    s += '<xf borderId="1" fillId="3" fontId="0" numFmtId="0" xfId="0" applyAlignment="1" applyBorder="1" applyFill="1"><alignment horizontal="center" vertical="center" wrapText="1"/></xf>'
    s += '<xf borderId="1" fillId="5" fontId="5" numFmtId="0" xfId="0" applyAlignment="1" applyBorder="1" applyFill="1" applyFont="1"><alignment horizontal="center" vertical="center" wrapText="1"/></xf>'
    s += '<xf borderId="1" fillId="0" fontId="0" numFmtId="164" xfId="0" applyAlignment="1" applyBorder="1" applyNumberFormat="1"><alignment horizontal="center" vertical="center" wrapText="1"/></xf>'
    s += '<xf borderId="0" fillId="0" fontId="3" numFmtId="0" xfId="0" applyAlignment="1" applyFont="1"><alignment vertical="center" wrapText="1"/></xf>'
    s += '<xf borderId="2" fillId="6" fontId="6" numFmtId="0" xfId="0" applyAlignment="1" applyBorder="1" applyFill="1" applyFont="1"><alignment horizontal="center" vertical="center" wrapText="1"/></xf>'
    s += '<xf borderId="1" fillId="7" fontId="3" numFmtId="164" xfId="0" applyAlignment="1" applyBorder="1" applyFill="1" applyFont="1" applyNumberFormat="1"><alignment horizontal="center" vertical="center" wrapText="1"/></xf>'
    s += '<xf borderId="1" fillId="0" fontId="3" numFmtId="0" xfId="0" applyAlignment="1" applyBorder="1" applyFont="1"><alignment vertical="center" wrapText="1"/></xf>'
    s += '<xf borderId="0" fillId="0" fontId="7" numFmtId="0" xfId="0" applyAlignment="1" applyFont="1"><alignment vertical="center" wrapText="1"/></xf>'
    s += '<xf borderId="1" fillId="0" fontId="0" numFmtId="165" xfId="0" applyAlignment="1" applyBorder="1" applyNumberFormat="1"><alignment horizontal="center" vertical="center" wrapText="1"/></xf>'
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



def build_kalkulator_sheet(ss_map):
    """Sheet 1: KALKULATOR - Interactive lookup/calculator sheet.
    
    Input cells:
    - B3: Gender (Pria/Wanita)
    - B4: Usia (Age, 0-85)
    - B5: Premi (Premium amount)
    - B6: Plan (A/B/C)
    - B7: Mata Uang (IDR/USD)
    - B8: Masa Pembayaran (Single/5/10)
    - B9: Masa Pertanggungan (Coverage period)
    
    Results section uses formulas referencing DATA sheets.
    """
    sb = SheetBuilder(ss_map)
    sb.set_col_width(1, 28)
    sb.set_col_width(2, 22)
    sb.set_col_width(3, 22)
    sb.set_col_width(4, 25)
    sb.set_col_width(5, 25)
    sb.set_col_width(6, 20)

    r = 1
    # Title
    sb.add_row(r, [(0, 'KALKULATOR MDWA - Manulife Dynamic Wealth Assurance', 5)])
    sb.add_merge('A1:F1')
    r += 2

    # === INPUT SECTION ===
    sb.add_row(r, [(0, 'INPUT DATA NASABAH', 7)])
    r += 1

    # Gender
    sb.add_row(r, [(0, 'Jenis Kelamin:', 10), (1, 'Pria', 8)])
    sb.add_data_validation('B4', '"Pria,Wanita"')
    r += 1

    # Age
    sb.add_row(r, [(0, 'Usia Masuk (tahun):', 10), (1, 30, 8, 'n')])
    r += 1

    # Premium
    sb.add_row(r, [(0, 'Premi (nominal):', 10), (1, 100000000, 8, 'n')])
    r += 1

    # Plan
    sb.add_row(r, [(0, 'Plan:', 10), (1, 'A', 8)])
    sb.add_data_validation('B7', '"A,B,C"')
    r += 1

    # Currency
    sb.add_row(r, [(0, 'Mata Uang:', 10), (1, 'IDR', 8)])
    sb.add_data_validation('B8', '"IDR,USD"')
    r += 1

    # Payment Term
    sb.add_row(r, [(0, 'Masa Pembayaran Premi:', 10), (1, 'Single', 8)])
    sb.add_data_validation('B9', '"Single,5,10"')
    r += 1

    # Coverage Period
    sb.add_row(r, [(0, 'Masa Pertanggungan (tahun):', 10), (1, 20, 8, 'n')])
    sb.add_data_validation('B10', '"5,6,7,8,9,10,11,12,13,14,15,20,30"')
    r += 1

    # Mode Pembayaran
    sb.add_row(r, [(0, 'Mode Pembayaran:', 10), (1, 'Tahunan', 8)])
    sb.add_data_validation('B11', '"Tahunan,Semesteran,Triwulanan,Bulanan"')
    r += 2

    # === ELIGIBILITY CHECK ===
    r = 13
    sb.add_row(r, [(0, 'HASIL PERHITUNGAN', 5)])
    sb.add_merge('A13:F13')
    r += 1

    # Eligibility formula
    # Single: age 0-85, Regular (5/10): age 0-70
    elig_formula = 'IF(B9="Single",IF(B5<=85,"ELIGIBLE","TIDAK ELIGIBLE - Usia maks 85 tahun"),IF(B5<=70,"ELIGIBLE","TIDAK ELIGIBLE - Usia maks 70 tahun untuk PPP 5/10"))'
    sb.add_row(r, [(0, 'Kelayakan Usia:', 10), (1, elig_formula, 9, 'f')])
    r += 1

    # Banding determination
    # Single IDR: <500M = B1, >=500M = B2; Regular IDR: <100M = B1, >=100M = B2
    # Single USD: <50000 = B1, >=50000 = B2; Regular USD: <10000 = B1, >=10000 = B2
    banding_formula = (
        'IF(B8="IDR",'
        'IF(B9="Single",IF(B6>=500000000,"Banding 2","Banding 1"),IF(B6>=100000000,"Banding 2","Banding 1")),'
        'IF(B9="Single",IF(B6>=50000,"Banding 2","Banding 1"),IF(B6>=10000,"Banding 2","Banding 1")))'
    )
    sb.add_row(r, [(0, 'Level Banding:', 10), (1, banding_formula, 9, 'f')])
    r += 1

    # Min premium check
    min_prem_formula = (
        'IF(B8="IDR",'
        'IF(B9="Single",IF(B6>=50000000,"OK - Min IDR 50 Juta","KURANG - Min IDR 50 Juta"),IF(B6>=24000000,"OK - Min IDR 24 Juta","KURANG - Min IDR 24 Juta")),'
        'IF(B9="Single",IF(B6>=5000,"OK - Min USD 5,000","KURANG - Min USD 5,000"),IF(B6>=2400,"OK - Min USD 2,400","KURANG - Min USD 2,400")))'
    )
    sb.add_row(r, [(0, 'Cek Minimum Premi:', 10), (1, min_prem_formula, 9, 'f')])
    r += 1

    # Coverage period validation
    cov_valid_formula = (
        'IF(B7="A",IF(OR(B10=20,B10=30),"VALID","Plan A hanya 20/30 tahun"),'
        'IF(B7="B",IF(OR(B10=5,B10=6,B10=7,B10=8,B10=9,B10=10,B10=11,B10=12,B10=13,B10=14,B10=15,B10=20,B10=30),"VALID","Plan B: 5-15/20/30 tahun"),'
        'IF(OR(B10=10,B10=15,B10=20,B10=30),"VALID","Plan C hanya 10/15/20/30 tahun")))'
    )
    sb.add_row(r, [(0, 'Validasi Masa Pertanggungan:', 10), (1, cov_valid_formula, 9, 'f')])
    r += 1

    # Payment term validation for coverage
    ppp_valid_formula = (
        'IF(B9="Single","VALID",'
        'IF(B9="5",IF(B10>5,"VALID","Coverage harus > PPP"),'
        'IF(B10>10,"VALID","Coverage harus > PPP")))'
    )
    sb.add_row(r, [(0, 'Validasi PPP vs Coverage:', 10), (1, ppp_valid_formula, 9, 'f')])
    r += 2

    # === BENEFIT CALCULATION ===
    r = 20
    sb.add_row(r, [(0, 'PROYEKSI MANFAAT', 5)])
    sb.add_merge('A20:F20')
    r += 1

    # Benefit percentage lookup using INDEX/MATCH from DATA sheets
    # The formula references DATA_PLAN_A, DATA_PLAN_B, DATA_PLAN_C sheets
    # Key in data sheets: Column A = coverage period, Column B = payment term label
    # We use a combined lookup: MATCH(B9&"_"&B10, key_column, 0)
    
    # Tahapan (Annual Payout) percentage
    tahapan_pct_formula = (
        'IF(B7="A",'
        'INDEX(DATA_PLAN_A!C:C,MATCH(B9&"_"&B10,DATA_PLAN_A!A:A,0)),'
        'IF(B7="C",'
        'INDEX(DATA_PLAN_C!C:C,MATCH(B9&"_"&B10,DATA_PLAN_C!A:A,0)),'
        '0))'
    )
    sb.add_row(r, [(0, 'Persentase Tahapan Tahunan:', 10), (1, tahapan_pct_formula, 9, 'f'),
                   (2, '(% dari Premi)', 11)])
    r += 1

    # Tahapan amount
    tahapan_amt_formula = (
        'IF(B7="B",0,B21*B6/100)'
    )
    sb.add_row(r, [(0, 'Nominal Tahapan/Tahun:', 10), (1, tahapan_amt_formula, 9, 'f'),
                   (2, 'IDR/USD per tahun', 11)])
    r += 1

    # Maturity percentage
    maturity_pct_formula = (
        'IF(B7="B",'
        'INDEX(DATA_PLAN_B!C:C,MATCH(B9&"_"&B10,DATA_PLAN_B!A:A,0)),'
        'IF(B7="C",'
        'INDEX(DATA_PLAN_C!D:D,MATCH(B9&"_"&B10,DATA_PLAN_C!A:A,0)),'
        '0))'
    )
    sb.add_row(r, [(0, 'Persentase Manfaat Jatuh Tempo:', 10), (1, maturity_pct_formula, 9, 'f'),
                   (2, '(% dari Premi)', 11)])
    r += 1

    # Maturity amount
    maturity_amt_formula = 'B23*B6/100'
    sb.add_row(r, [(0, 'Nominal Manfaat Jatuh Tempo:', 10), (1, maturity_amt_formula, 9, 'f')])
    r += 1

    # Banding 2 Extra percentage
    extra_pct_formula = (
        'IF(B15="Banding 1",0,'
        'IF(B7="A",'
        'INDEX(DATA_PLAN_A!D:D,MATCH(B9&"_"&B10,DATA_PLAN_A!A:A,0)),'
        'IF(B7="B",'
        'INDEX(DATA_PLAN_B!D:D,MATCH(B9&"_"&B10,DATA_PLAN_B!A:A,0)),'
        'INDEX(DATA_PLAN_C!E:E,MATCH(B9&"_"&B10,DATA_PLAN_C!A:A,0)))))'
    )
    sb.add_row(r, [(0, 'Banding 2 Extra (%):', 10), (1, extra_pct_formula, 9, 'f')])
    r += 1

    # Banding 2 Extra amount
    extra_amt_formula = 'B25*B6/100'
    sb.add_row(r, [(0, 'Banding 2 Extra (nominal):', 10), (1, extra_amt_formula, 9, 'f')])
    r += 2

    # Payout start year
    r = 28
    payout_start_formula = (
        'IF(B7="B","N/A (lump sum di akhir)",'
        'IF(B9="Single","Tahun ke-6",'
        'IF(B9="5","Tahun ke-10","Tahun ke-15")))'
    )
    sb.add_row(r, [(0, 'Mulai Pembayaran Tahapan:', 10), (1, payout_start_formula, 9, 'f')])
    r += 1

    # Number of payout years
    payout_years_formula = (
        'IF(B7="B",0,'
        'IF(B9="Single",B10-5,'
        'IF(B9="5",B10-9,B10-14)))'
    )
    sb.add_row(r, [(0, 'Jumlah Tahun Pembayaran:', 10), (1, payout_years_formula, 9, 'f')])
    r += 1

    # Total tahapan
    total_tahapan_formula = 'IF(B7="B",0,B22*B29)'
    sb.add_row(r, [(0, 'Total Tahapan (seluruh periode):', 10), (1, total_tahapan_formula, 9, 'f')])
    r += 1

    # Total benefit
    total_benefit_formula = 'B30+B24+B26'
    sb.add_row(r, [(0, 'TOTAL MANFAAT:', 7), (1, total_benefit_formula, 9, 'f')])
    r += 2

    # === ADB & TPD ===
    r = 33
    sb.add_row(r, [(0, 'MANFAAT TAMBAHAN', 5)])
    sb.add_merge('A33:F33')
    r += 1

    # ADB & TPD = 25% of annual basic premium per year
    adb_formula = (
        'IF(B9="Single",B6*0.25/B10,B6*0.25)'
    )
    sb.add_row(r, [(0, 'ADB & TPD per tahun:', 10), (1, adb_formula, 9, 'f'),
                   (2, '25% of Annual Basic Premium', 11)])
    r += 1

    total_adb_formula = 'B34*B10'
    sb.add_row(r, [(0, 'Total ADB & TPD (full coverage):', 10), (1, total_adb_formula, 9, 'f')])
    r += 2

    # === MODE PEMBAYARAN ===
    r = 37
    sb.add_row(r, [(0, 'KONVERSI MODE PEMBAYARAN', 5)])
    sb.add_merge('A37:F37')
    r += 1

    # Mode factor formulas
    mode_tahunan_formula = 'B6'
    sb.add_row(r, [(0, 'Premi Tahunan:', 10), (1, mode_tahunan_formula, 9, 'f')])
    r += 1

    mode_semester_formula = 'IF(B8="IDR",B6*0.525,B6*0.5125)'
    sb.add_row(r, [(0, 'Premi Semesteran:', 10), (1, mode_semester_formula, 9, 'f'),
                   (2, 'IDR 52.5% / USD 51.25%', 11)])
    r += 1

    mode_triwulan_formula = 'IF(B8="IDR",B6*0.275,B6*0.265)'
    sb.add_row(r, [(0, 'Premi Triwulanan:', 10), (1, mode_triwulan_formula, 9, 'f'),
                   (2, 'IDR 27.5% / USD 26.5%', 11)])
    r += 1

    mode_bulanan_formula = 'IF(B8="IDR",B6*0.095,B6*0.09)'
    sb.add_row(r, [(0, 'Premi Bulanan:', 10), (1, mode_bulanan_formula, 9, 'f'),
                   (2, 'IDR 9.5% / USD 9%', 11)])
    r += 2

    # === NOTES ===
    r = 43
    sb.add_row(r, [(0, 'CATATAN:', 7)])
    r += 1
    sb.add_row(r, [(0, '- Ubah nilai di sel berwarna KUNING untuk melihat hasil otomatis', 11)])
    r += 1
    sb.add_row(r, [(0, '- Plan A (Anuitas): Tahapan tahunan, coverage 20/30 tahun', 11)])
    r += 1
    sb.add_row(r, [(0, '- Plan B (Beasiswa): Lump sum jatuh tempo, coverage 5-15/20/30 tahun', 11)])
    r += 1
    sb.add_row(r, [(0, '- Plan C (Combo): Tahapan + jatuh tempo, coverage 10/15/20/30 tahun', 11)])
    r += 1
    sb.add_row(r, [(0, '- Lihat sheet DATA_PLAN_A/B/C untuk tabel referensi lengkap', 11)])
    r += 1
    sb.add_row(r, [(0, '- Usia masuk: Single 30 hari-85 thn, Regular 30 hari-70 thn', 11)])
    r += 1
    sb.add_row(r, [(0, '- Gender mempengaruhi underwriting namun TIDAK mempengaruhi benefit MDWA (unisex rate)', 11)])

    return sb.build()



def build_data_plan_a_sheet(ss_map):
    """Sheet 2: DATA_PLAN_A - Reference data for Plan A (Anuitas) benefits.
    
    Layout:
    Column A: Key (PaymentTerm_CoveragePeriod) e.g. "Single_20", "5_30", "10_20"
    Column B: Payment Term label
    Column C: Tahapan % (Banding 1 IDR)
    Column D: Banding 2 Extra % (IDR)
    Column E: Tahapan % (Banding 1 USD)
    Column F: Banding 2 Extra % (USD)
    Column G: Payout Start Year
    """
    sb = SheetBuilder(ss_map)
    sb.set_col_width(1, 18)
    sb.set_col_width(2, 18)
    sb.set_col_width(3, 20)
    sb.set_col_width(4, 20)
    sb.set_col_width(5, 20)
    sb.set_col_width(6, 20)
    sb.set_col_width(7, 18)

    r = 1
    sb.add_row(r, [(0, 'DATA PLAN A - ANUITAS (Tabel Referensi)', 5)])
    sb.add_merge('A1:G1')
    r += 2

    # Header
    sb.add_row(r, [
        (0, 'Key', 1), (1, 'Masa Bayar', 1), (2, 'Tahapan % IDR (B1)', 1),
        (3, 'Extra % IDR (B2)', 1), (4, 'Tahapan % USD (B1)', 1),
        (5, 'Extra % USD (B2)', 1), (6, 'Mulai Bayar', 1)
    ])
    r += 1

    # Plan A data: Coverage 20yr and 30yr for Single, 5yr, 10yr
    plan_a_data = [
        # (key, term_label, tahapan_idr, extra_idr, tahapan_usd, extra_usd, start_yr)
        ('Single_20', 'Single', 9.5, 6, 8.25, 3, 6),
        ('Single_30', 'Single', 7, 9, 5.5, 4.5, 6),
        ('5_20', '5 Tahun', 65, 30, 55, 15, 10),
        ('5_30', '5 Tahun', 40, 45, 32.5, 22.5, 10),
        ('10_20', '10 Tahun', 235, 60, 205, 30, 15),
        ('10_30', '10 Tahun', 105, 90, 85, 45, 15),
    ]

    for key, term, t_idr, e_idr, t_usd, e_usd, start in plan_a_data:
        sb.add_row(r, [
            (0, key, 2), (1, term, 2), (2, t_idr, 6, 'n'),
            (3, e_idr, 6, 'n'), (4, t_usd, 6, 'n'),
            (5, e_usd, 6, 'n'), (6, start, 6, 'n')
        ])
        r += 1

    r += 2
    # Additional info
    sb.add_row(r, [(0, 'JADWAL PEMBAYARAN TAHAPAN', 7)])
    r += 1
    sb.add_row(r, [(0, 'Single: Mulai akhir tahun ke-6 s/d akhir masa pertanggungan', 0)])
    r += 1
    sb.add_row(r, [(0, 'PPP 5: Mulai akhir tahun ke-10 s/d akhir masa pertanggungan', 0)])
    r += 1
    sb.add_row(r, [(0, 'PPP 10: Mulai akhir tahun ke-15 s/d akhir masa pertanggungan', 0)])
    r += 2
    sb.add_row(r, [(0, 'CATATAN:', 7)])
    r += 1
    sb.add_row(r, [(0, 'Tahapan = % x Premi, dibayar setiap tahun dari start sampai akhir coverage', 0)])
    r += 1
    sb.add_row(r, [(0, 'Banding 2 Extra = tambahan % untuk premi >= IDR 500M (Single) atau >= IDR 100M (Regular)', 0)])

    return sb.build()



def build_data_plan_b_sheet(ss_map):
    """Sheet 3: DATA_PLAN_B - Reference data for Plan B (Beasiswa) benefits.
    
    Column A: Key (PaymentTerm_CoveragePeriod)
    Column B: Payment Term label
    Column C: Maturity % (Banding 1 IDR)
    Column D: Banding 2 Extra % (IDR)
    Column E: Maturity % (Banding 1 USD)
    Column F: Banding 2 Extra % (USD)
    """
    sb = SheetBuilder(ss_map)
    sb.set_col_width(1, 18)
    sb.set_col_width(2, 18)
    sb.set_col_width(3, 22)
    sb.set_col_width(4, 22)
    sb.set_col_width(5, 22)
    sb.set_col_width(6, 22)

    r = 1
    sb.add_row(r, [(0, 'DATA PLAN B - BEASISWA (Tabel Referensi)', 5)])
    sb.add_merge('A1:F1')
    r += 2

    # Header
    sb.add_row(r, [
        (0, 'Key', 1), (1, 'Masa Bayar', 1), (2, 'Maturity % IDR (B1)', 1),
        (3, 'Extra % IDR (B2)', 1), (4, 'Maturity % USD (B1)', 1),
        (5, 'Extra % USD (B2)', 1)
    ])
    r += 1

    # Plan B data
    plan_b_data = [
        # Single Premium
        ('Single_5', 'Single', 110, 1.5, 105, 0.75),
        ('Single_6', 'Single', 115, 1.8, 106, 0.9),
        ('Single_7', 'Single', 120, 2.1, 108, 1.05),
        ('Single_8', 'Single', 125, 2.4, 110, 1.2),
        ('Single_9', 'Single', 130, 2.7, 112, 1.35),
        ('Single_10', 'Single', 135, 3, 114, 1.5),
        ('Single_11', 'Single', 140, 3.3, 116, 1.65),
        ('Single_12', 'Single', 145, 3.6, 120, 1.8),
        ('Single_13', 'Single', 150, 3.9, 125, 1.95),
        ('Single_14', 'Single', 155, 4.2, 130, 2.1),
        ('Single_15', 'Single', 160, 4.5, 135, 2.25),
        ('Single_20', 'Single', 200, 6, 160, 3),
        ('Single_30', 'Single', 325, 9, 200, 4.5),
        # PPP 5
        ('5_10', '5 Tahun', 600, 15, 535, 7.5),
        ('5_11', '5 Tahun', 625, 16.5, 545, 8.25),
        ('5_12', '5 Tahun', 650, 18, 555, 9),
        ('5_13', '5 Tahun', 675, 19.5, 565, 9.75),
        ('5_14', '5 Tahun', 700, 21, 575, 10.5),
        ('5_15', '5 Tahun', 725, 22.5, 600, 11.25),
        ('5_20', '5 Tahun', 950, 30, 700, 15),
        ('5_30', '5 Tahun', 1625, 45, 950, 22.5),
        # PPP 10
        ('10_15', '10 Tahun', 1350, 45, 1175, 22.5),
        ('10_20', '10 Tahun', 1800, 60, 1350, 30),
        ('10_30', '10 Tahun', 3250, 90, 1800, 45),
    ]

    for key, term, mat_idr, ext_idr, mat_usd, ext_usd in plan_b_data:
        sb.add_row(r, [
            (0, key, 2), (1, term, 2), (2, mat_idr, 6, 'n'),
            (3, ext_idr, 6, 'n'), (4, mat_usd, 6, 'n'),
            (5, ext_usd, 6, 'n')
        ])
        r += 1

    return sb.build()



def build_data_plan_c_sheet(ss_map):
    """Sheet 4: DATA_PLAN_C - Reference data for Plan C (Combo) benefits.
    
    Column A: Key (PaymentTerm_CoveragePeriod)
    Column B: Payment Term label
    Column C: Tahapan % (Banding 1 IDR)
    Column D: Maturity % (Banding 1 IDR)
    Column E: Banding 2 Extra Maturity % (IDR)
    Column F: Tahapan % (Banding 1 USD)
    Column G: Maturity % (Banding 1 USD)
    Column H: Banding 2 Extra Maturity % (USD)
    """
    sb = SheetBuilder(ss_map)
    for i in range(1, 9):
        sb.set_col_width(i, 18)

    r = 1
    sb.add_row(r, [(0, 'DATA PLAN C - COMBO (Tabel Referensi)', 5)])
    sb.add_merge('A1:H1')
    r += 2

    # Header
    sb.add_row(r, [
        (0, 'Key', 1), (1, 'Masa Bayar', 1), (2, 'Tahapan % IDR', 1),
        (3, 'Maturity % IDR', 1), (4, 'Extra Mat % IDR', 1),
        (5, 'Tahapan % USD', 1), (6, 'Maturity % USD', 1),
        (7, 'Extra Mat % USD', 1)
    ])
    r += 1

    # Plan C data
    plan_c_data = [
        # Single Premium
        ('Single_10', 'Single', 2, 105, 3, 1, 100, 1),
        ('Single_15', 'Single', 2, 120, 5, 1, 110, 2.25),
        ('Single_20', 'Single', 2, 135, 6, 1, 120, 3),
        ('Single_30', 'Single', 2, 170, 9, 1, 140, 5),
        # PPP 5
        ('5_10', '5 Tahun', 13.5, 510, 15, 5, 500, 8),
        ('5_15', '5 Tahun', 17.5, 510, 23, 8, 500, 11.25),
        ('5_20', '5 Tahun', 20, 510, 30, 11, 500, 15),
        ('5_30', '5 Tahun', 20, 510, 45, 11, 500, 23),
        # PPP 10
        ('10_15', '10 Tahun', 40, 1000, 45, 20, 1000, 23),
        ('10_20', '10 Tahun', 40, 1100, 60, 20, 1100, 30),
        ('10_30', '10 Tahun', 40, 1200, 90, 20, 1150, 45),
    ]

    for key, term, t_idr, m_idr, e_idr, t_usd, m_usd, e_usd in plan_c_data:
        sb.add_row(r, [
            (0, key, 2), (1, term, 2), (2, t_idr, 6, 'n'),
            (3, m_idr, 6, 'n'), (4, e_idr, 6, 'n'),
            (5, t_usd, 6, 'n'), (6, m_usd, 6, 'n'),
            (7, e_usd, 6, 'n')
        ])
        r += 1

    r += 2
    sb.add_row(r, [(0, 'JADWAL PEMBAYARAN TAHAPAN PLAN C', 7)])
    r += 1
    sb.add_row(r, [(0, 'Single: Mulai akhir tahun ke-1 s/d akhir masa pertanggungan', 0)])
    r += 1
    sb.add_row(r, [(0, 'PPP 5: Mulai akhir tahun ke-5 s/d akhir masa pertanggungan', 0)])
    r += 1
    sb.add_row(r, [(0, 'PPP 10: Mulai akhir tahun ke-10 s/d akhir masa pertanggungan', 0)])

    return sb.build()



def build_komisi_sheet(ss_map):
    """Sheet 5: KOMISI - Commission Structure."""
    sb = SheetBuilder(ss_map)
    sb.set_col_width(1, 20)
    sb.set_col_width(2, 18)
    sb.set_col_width(3, 15)
    sb.set_col_width(4, 15)

    r = 1
    sb.add_row(r, [(0, 'STRUKTUR KOMISI MDWA', 5)])
    sb.add_merge('A1:D1')
    r += 2

    # Plan A Commission
    sb.add_row(r, [(0, 'PLAN A - ANUITAS', 7)])
    r += 1
    sb.add_row(r, [(0, 'Term', 1), (1, 'Coverage', 1), (2, 'Year 1', 1), (3, 'Year 2', 1)])
    r += 1
    plan_a_data = [
        ('Single', '20yr', '4%', '-'),
        ('Single', '30yr', '4%', '-'),
        ('5 Pay', '20yr', '5%', '5%'),
        ('5 Pay', '30yr', '5%', '5%'),
        ('10 Pay', '20yr', '15%', '10%'),
        ('10 Pay', '30yr', '15%', '10%'),
    ]
    for term, cov, y1, y2 in plan_a_data:
        sb.add_row(r, [(0, term, 2), (1, cov, 2), (2, y1, 2), (3, y2, 2)])
        r += 1
    r += 1

    # Plan B Commission
    sb.add_row(r, [(0, 'PLAN B - BEASISWA', 7)])
    r += 1
    sb.add_row(r, [(0, 'Term', 1), (1, 'Coverage', 1), (2, 'Year 1', 1), (3, 'Year 2', 1)])
    r += 1
    plan_b_data = [
        ('Single', '5-10yr', '2%', '-'),
        ('Single', '11-15yr', '5%', '-'),
        ('Single', '20yr', '5%', '-'),
        ('Single', '30yr', '6%', '-'),
        ('5 Pay', '5-10yr', '7.5%', '5%'),
        ('5 Pay', '11-15yr', '7.5%', '5%'),
        ('5 Pay', '20yr', '10%', '5%'),
        ('5 Pay', '30yr', '10%', '5%'),
        ('10 Pay', 'all', '15%', '10%'),
    ]
    for term, cov, y1, y2 in plan_b_data:
        sb.add_row(r, [(0, term, 2), (1, cov, 2), (2, y1, 2), (3, y2, 2)])
        r += 1
    r += 1

    # Plan C Commission
    sb.add_row(r, [(0, 'PLAN C - COMBO', 7)])
    r += 1
    sb.add_row(r, [(0, 'Term', 1), (1, 'Coverage', 1), (2, 'Year 1', 1), (3, 'Year 2', 1)])
    r += 1
    plan_c_data = [
        ('Single', '10yr', '2%', '-'),
        ('Single', '15yr', '2%', '-'),
        ('Single', '20yr', '4%', '-'),
        ('Single', '30yr', '6%', '-'),
        ('5 Pay', '10yr', '5%', '5%'),
        ('5 Pay', '15yr', '7.5%', '5%'),
        ('5 Pay', '20yr', '10%', '5%'),
        ('5 Pay', '30yr', '12.5%', '5%'),
        ('10 Pay', 'all', '15%', '10%'),
    ]
    for term, cov, y1, y2 in plan_c_data:
        sb.add_row(r, [(0, term, 2), (1, cov, 2), (2, y1, 2), (3, y2, 2)])
        r += 1

    return sb.build()



def build_perbandingan_sheet(ss_map):
    """Sheet 6: PERBANDINGAN - Comparison vs Competitors."""
    sb = SheetBuilder(ss_map)
    for i in range(1, 8):
        sb.set_col_width(i, 18)

    r = 1
    sb.add_row(r, [(0, 'PERBANDINGAN vs KOMPETITOR', 5)])
    sb.add_merge('A1:G1')
    r += 2

    sb.add_row(r, [(0, 'PPP 5 Tahun, Band 2, Premi Tahunan IDR 100 Juta', 7)])
    r += 2

    # Header
    sb.add_row(r, [(0, 'Feature', 1), (1, 'MDWA 15yr', 1), (2, 'MDWA 20yr', 1),
                   (3, 'MSP 15yr', 1), (4, 'MSP 20yr', 1), (5, 'FI 15yr', 1), (6, 'FI 20yr', 1)])
    r += 1
    sb.add_row(r, [(0, 'Tahapan/Tahun', 2), (1, '17.5M x 11', 2), (2, '20M x 16', 2),
                   (3, '17.5M x 11', 2), (4, '20M x 16', 2), (5, '20M x 10', 2), (6, '20M x 15', 2)])
    r += 1
    sb.add_row(r, [(0, 'Total Tahapan', 2), (1, '192.5M', 2), (2, '320M', 2),
                   (3, '192.5M', 2), (4, '320M', 2), (5, '200M', 2), (6, '300M', 2)])
    r += 1
    sb.add_row(r, [(0, 'Maturity', 2), (1, '532.5M', 2), (2, '540M', 2),
                   (3, '500M', 2), (4, '500M', 2), (5, '500M', 2), (6, '525M', 2)])
    r += 1
    sb.add_row(r, [(0, 'Total Benefit', 4), (1, '725M', 4), (2, '860M', 4),
                   (3, '692.5M', 4), (4, '820M', 4), (5, '700M', 4), (6, '825M', 4)])
    r += 1
    sb.add_row(r, [(0, 'Rasio Benefit', 4), (1, '145%', 4), (2, '172%', 4),
                   (3, '139%', 4), (4, '164%', 4), (5, '140%', 4), (6, '165%', 4)])
    r += 2

    sb.add_row(r, [(0, 'Catatan:', 7)])
    r += 1
    sb.add_row(r, [(0, 'MSP = Manulife Saver Plus (produk lama)', 0)])
    r += 1
    sb.add_row(r, [(0, 'FI = Flexi Income (kompetitor)', 0)])
    r += 1
    sb.add_row(r, [(0, 'Rasio Benefit = Total Benefit / Total Premi (5 tahun x IDR 100 Juta = IDR 500 Juta)', 0)])
    r += 2
    sb.add_row(r, [(0, 'MDWA memberikan rasio benefit tertinggi di kedua masa pertanggungan.', 7)])

    return sb.build()



def build_overview_sheet(ss_map):
    """Sheet 7: OVERVIEW - Product Summary."""
    sb = SheetBuilder(ss_map)
    sb.set_col_width(1, 25)
    sb.set_col_width(2, 30)
    sb.set_col_width(3, 30)
    sb.set_col_width(4, 30)

    r = 1
    sb.add_row(r, [(0, 'MANULIFE DYNAMIC WEALTH ASSURANCE (MDWA)', 5)])
    sb.add_merge('A1:D1')
    r += 1
    sb.add_row(r, [(0, 'Perbandingan 3 Plan MDWA', 3)])
    r += 2

    # Plan comparison header
    sb.add_row(r, [(0, 'Fitur', 1), (1, 'Plan A - Anuitas', 1), (2, 'Plan B - Beasiswa', 1), (3, 'Plan C - Combo', 1)])
    r += 1
    sb.add_row(r, [(0, 'Konsep', 2), (1, 'Arus kas tahunan stabil untuk perencanaan pensiun', 2), (2, 'Lump sum di akhir masa pertanggungan untuk dana pendidikan', 2), (3, 'Pembayaran tahunan + manfaat jatuh tempo untuk tabungan fleksibel', 2)])
    r += 1
    sb.add_row(r, [(0, 'Masa Pertanggungan', 2), (1, '20 / 30 tahun', 2), (2, '5-15 / 20 / 30 tahun', 2), (3, '10 / 15 / 20 / 30 tahun', 2)])
    r += 1
    sb.add_row(r, [(0, 'Manfaat Utama', 2), (1, 'Tahapan tahunan (annual payout)', 2), (2, 'Lump sum saat jatuh tempo (maturity)', 2), (3, 'Tahapan tahunan + lump sum jatuh tempo', 2)])
    r += 2

    # Common specifications
    sb.add_row(r, [(0, 'SPESIFIKASI UMUM', 5)])
    sb.add_merge('A' + str(r) + ':D' + str(r))
    r += 1
    specs = [
        ('Masa Pembayaran Premi', 'Premi Sekaligus (Single Premium), PPP 5 tahun, PPP 10 tahun'),
        ('Mode Pembayaran', 'Tahunan (100%), Semesteran (52.5%/51.25%), Tiga Bulanan (27.5%/26.5%), Bulanan (9.5%/9%)'),
        ('Usia Masuk (Single)', '30 hari - 85 tahun'),
        ('Usia Masuk (Regular)', '30 hari - 70 tahun'),
        ('Mata Uang', 'IDR dan USD'),
        ('Underwriting', 'Guaranteed Issuance Offering'),
        ('Riders', 'Advanced Life Protector Plus, Manulife Payor Benefit Plus, Manulife Waiver of Premium Plus'),
        ('Min. Premi Single', 'IDR 50 Juta / USD 5,000'),
        ('Min. Premi Regular', 'IDR 24 Juta / USD 2,400'),
        ('ADB & TPD', '25% of Annual Basic Premium, dibayar tahunan sampai akhir masa pertanggungan'),
    ]
    for label, val in specs:
        sb.add_row(r, [(0, label, 4), (1, val, 2)])
        sb.add_merge('B' + str(r) + ':D' + str(r))
        r += 1

    r += 1
    sb.add_row(r, [(0, 'PREMIUM BANDING', 5)])
    sb.add_merge('A' + str(r) + ':D' + str(r))
    r += 1
    sb.add_row(r, [(0, 'Banding', 1), (1, 'Single Premium', 1), (2, 'Regular Premium', 1)])
    r += 1
    sb.add_row(r, [(0, 'Banding 1 (IDR)', 2), (1, 'IDR 50M - <500M', 2), (2, 'IDR 24M - <100M', 2)])
    r += 1
    sb.add_row(r, [(0, 'Banding 1 (USD)', 2), (1, 'USD 5,000 - <50,000', 2), (2, 'USD 2,400 - <10,000', 2)])
    r += 1
    sb.add_row(r, [(0, 'Banding 2 (IDR)', 2), (1, '>= IDR 500M', 2), (2, '>= IDR 100M', 2)])
    r += 1
    sb.add_row(r, [(0, 'Banding 2 (USD)', 2), (1, '>= USD 50,000', 2), (2, '>= USD 10,000', 2)])

    return sb.build()



def main():
    print('Generating Kalkulator_MDWA.xlsx (with interactive KALKULATOR sheet)...')

    # Collect all unique strings
    all_strings = []
    seen = set()

    def add_str(s):
        if s not in seen:
            seen.add(s)
            all_strings.append(s)

    # Gather strings from all sheets
    kalk_strs = [
        'KALKULATOR MDWA - Manulife Dynamic Wealth Assurance',
        'INPUT DATA NASABAH', 'Jenis Kelamin:', 'Pria', 'Wanita',
        'Usia Masuk (tahun):', 'Premi (nominal):', 'Plan:', 'A', 'B', 'C',
        'Mata Uang:', 'IDR', 'USD', 'Masa Pembayaran Premi:', 'Single',
        'Masa Pertanggungan (tahun):', 'Mode Pembayaran:', 'Tahunan',
        'HASIL PERHITUNGAN', 'Kelayakan Usia:', 'Level Banding:',
        'Cek Minimum Premi:', 'Validasi Masa Pertanggungan:',
        'Validasi PPP vs Coverage:',
        'PROYEKSI MANFAAT', 'Persentase Tahapan Tahunan:', '(% dari Premi)',
        'Nominal Tahapan/Tahun:', 'IDR/USD per tahun',
        'Persentase Manfaat Jatuh Tempo:', 'Nominal Manfaat Jatuh Tempo:',
        'Banding 2 Extra (%):', 'Banding 2 Extra (nominal):',
        'Mulai Pembayaran Tahapan:', 'Jumlah Tahun Pembayaran:',
        'Total Tahapan (seluruh periode):', 'TOTAL MANFAAT:',
        'MANFAAT TAMBAHAN', 'ADB & TPD per tahun:',
        '25% of Annual Basic Premium', 'Total ADB & TPD (full coverage):',
        'KONVERSI MODE PEMBAYARAN', 'Premi Tahunan:', 'Premi Semesteran:',
        'IDR 52.5% / USD 51.25%', 'Premi Triwulanan:', 'IDR 27.5% / USD 26.5%',
        'Premi Bulanan:', 'IDR 9.5% / USD 9%',
        'CATATAN:',
        '- Ubah nilai di sel berwarna KUNING untuk melihat hasil otomatis',
        '- Plan A (Anuitas): Tahapan tahunan, coverage 20/30 tahun',
        '- Plan B (Beasiswa): Lump sum jatuh tempo, coverage 5-15/20/30 tahun',
        '- Plan C (Combo): Tahapan + jatuh tempo, coverage 10/15/20/30 tahun',
        '- Lihat sheet DATA_PLAN_A/B/C untuk tabel referensi lengkap',
        '- Usia masuk: Single 30 hari-85 thn, Regular 30 hari-70 thn',
        '- Gender mempengaruhi underwriting namun TIDAK mempengaruhi benefit MDWA (unisex rate)',
    ]
    for s in kalk_strs:
        add_str(s)

    data_a_strs = [
        'DATA PLAN A - ANUITAS (Tabel Referensi)',
        'Key', 'Masa Bayar', 'Tahapan % IDR (B1)', 'Extra % IDR (B2)',
        'Tahapan % USD (B1)', 'Extra % USD (B2)', 'Mulai Bayar',
        'Single_20', 'Single_30', '5_20', '5_30', '10_20', '10_30',
        '5 Tahun', '10 Tahun',
        'JADWAL PEMBAYARAN TAHAPAN',
        'Single: Mulai akhir tahun ke-6 s/d akhir masa pertanggungan',
        'PPP 5: Mulai akhir tahun ke-10 s/d akhir masa pertanggungan',
        'PPP 10: Mulai akhir tahun ke-15 s/d akhir masa pertanggungan',
        'Tahapan = % x Premi, dibayar setiap tahun dari start sampai akhir coverage',
        'Banding 2 Extra = tambahan % untuk premi >= IDR 500M (Single) atau >= IDR 100M (Regular)',
    ]
    for s in data_a_strs:
        add_str(s)

    data_b_strs = [
        'DATA PLAN B - BEASISWA (Tabel Referensi)',
        'Maturity % IDR (B1)', 'Extra % IDR (B2)', 'Maturity % USD (B1)', 'Extra % USD (B2)',
        'Single_5', 'Single_6', 'Single_7', 'Single_8', 'Single_9', 'Single_10',
        'Single_11', 'Single_12', 'Single_13', 'Single_14', 'Single_15',
        '5_10', '5_11', '5_12', '5_13', '5_14', '5_15',
        '10_15',
    ]
    for s in data_b_strs:
        add_str(s)

    data_c_strs = [
        'DATA PLAN C - COMBO (Tabel Referensi)',
        'Tahapan % IDR', 'Maturity % IDR', 'Extra Mat % IDR',
        'Tahapan % USD', 'Maturity % USD', 'Extra Mat % USD',
        'Single_10', 'Single_15',
        '5_10', '5_15', '5_20', '5_30',
        '10_15', '10_20', '10_30',
        'JADWAL PEMBAYARAN TAHAPAN PLAN C',
        'Single: Mulai akhir tahun ke-1 s/d akhir masa pertanggungan',
        'PPP 5: Mulai akhir tahun ke-5 s/d akhir masa pertanggungan',
        'PPP 10: Mulai akhir tahun ke-10 s/d akhir masa pertanggungan',
    ]
    for s in data_c_strs:
        add_str(s)

    komisi_strs = [
        'STRUKTUR KOMISI MDWA',
        'PLAN A - ANUITAS', 'Term', 'Year 1', 'Year 2',
        '4%', '5%', '10%', '15%', '12.5%', '7.5%', '-',
        '20yr', '30yr', '5 Pay', '10 Pay',
        'PLAN B - BEASISWA',
        '5-10yr', '11-15yr', '2%', '6%', 'all',
        'PLAN C - COMBO', '10yr', '15yr',
    ]
    for s in komisi_strs:
        add_str(s)

    perb_strs = [
        'PERBANDINGAN vs KOMPETITOR',
        'PPP 5 Tahun, Band 2, Premi Tahunan IDR 100 Juta',
        'Feature', 'MDWA 15yr', 'MDWA 20yr', 'MSP 15yr', 'MSP 20yr', 'FI 15yr', 'FI 20yr',
        'Tahapan/Tahun', '17.5M x 11', '20M x 16', '20M x 10', '20M x 15',
        'Total Tahapan', '192.5M', '320M', '200M', '300M',
        'Maturity', '532.5M', '540M', '500M', '525M',
        'Total Benefit', '725M', '860M', '692.5M', '820M', '700M', '825M',
        'Rasio Benefit', '145%', '172%', '139%', '164%', '140%', '165%',
        'Catatan:',
        'MSP = Manulife Saver Plus (produk lama)',
        'FI = Flexi Income (kompetitor)',
        'Rasio Benefit = Total Benefit / Total Premi (5 tahun x IDR 100 Juta = IDR 500 Juta)',
        'MDWA memberikan rasio benefit tertinggi di kedua masa pertanggungan.',
    ]
    for s in perb_strs:
        add_str(s)

    overview_strs = [
        'MANULIFE DYNAMIC WEALTH ASSURANCE (MDWA)',
        'Perbandingan 3 Plan MDWA',
        'Fitur', 'Plan A - Anuitas', 'Plan B - Beasiswa', 'Plan C - Combo',
        'Konsep', 'Arus kas tahunan stabil untuk perencanaan pensiun',
        'Lump sum di akhir masa pertanggungan untuk dana pendidikan',
        'Pembayaran tahunan + manfaat jatuh tempo untuk tabungan fleksibel',
        'Masa Pertanggungan', '20 / 30 tahun', '5-15 / 20 / 30 tahun', '10 / 15 / 20 / 30 tahun',
        'Manfaat Utama', 'Tahapan tahunan (annual payout)',
        'Lump sum saat jatuh tempo (maturity)', 'Tahapan tahunan + lump sum jatuh tempo',
        'SPESIFIKASI UMUM',
        'Masa Pembayaran Premi', 'Premi Sekaligus (Single Premium), PPP 5 tahun, PPP 10 tahun',
        'Mode Pembayaran', 'Tahunan (100%), Semesteran (52.5%/51.25%), Tiga Bulanan (27.5%/26.5%), Bulanan (9.5%/9%)',
        'Usia Masuk (Single)', '30 hari - 85 tahun',
        'Usia Masuk (Regular)', '30 hari - 70 tahun',
        'Mata Uang', 'IDR dan USD',
        'Underwriting', 'Guaranteed Issuance Offering',
        'Riders', 'Advanced Life Protector Plus, Manulife Payor Benefit Plus, Manulife Waiver of Premium Plus',
        'Min. Premi Single', 'IDR 50 Juta / USD 5,000',
        'Min. Premi Regular', 'IDR 24 Juta / USD 2,400',
        'ADB & TPD', '25% of Annual Basic Premium, dibayar tahunan sampai akhir masa pertanggungan',
        'PREMIUM BANDING',
        'Banding', 'Single Premium', 'Regular Premium',
        'Banding 1 (IDR)', 'IDR 50M - <500M', 'IDR 24M - <100M',
        'Banding 1 (USD)', 'USD 5,000 - <50,000', 'USD 2,400 - <10,000',
        'Banding 2 (IDR)', '>= IDR 500M', '>= IDR 100M',
        'Banding 2 (USD)', '>= USD 50,000', '>= USD 10,000',
    ]
    for s in overview_strs:
        add_str(s)

    ss_map = {s: i for i, s in enumerate(all_strings)}

    print('  Building sheets...')
    sheet1_xml = build_kalkulator_sheet(ss_map)
    sheet2_xml = build_data_plan_a_sheet(ss_map)
    sheet3_xml = build_data_plan_b_sheet(ss_map)
    sheet4_xml = build_data_plan_c_sheet(ss_map)
    sheet5_xml = build_komisi_sheet(ss_map)
    sheet6_xml = build_perbandingan_sheet(ss_map)
    sheet7_xml = build_overview_sheet(ss_map)

    shared_strings_xml = build_shared_strings(all_strings)
    styles_xml = build_styles_xml()
    theme_xml = build_theme_xml()

    sheet_names = [
        'KALKULATOR', 'DATA_PLAN_A', 'DATA_PLAN_B', 'DATA_PLAN_C',
        'KOMISI', 'PERBANDINGAN', 'OVERVIEW'
    ]

    # Content Types
    ct = chr(60) + '?xml version="1.0" encoding="UTF-8" standalone="yes"?' + chr(62)
    ct += '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
    ct += '<Default ContentType="application/xml" Extension="xml"/>'
    ct += '<Default ContentType="application/vnd.openxmlformats-package.relationships+xml" Extension="rels"/>'
    for i in range(1, 8):
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
    for i in range(1, 8):
        wb_rels += '<Relationship Id="rId' + str(i + 3) + '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet' + str(i) + '.xml"/>'
    wb_rels += '</Relationships>'

    print('  Packaging xlsx...')
    with zipfile.ZipFile(OUTPUT_XLSX, 'w', zipfile.ZIP_DEFLATED) as zf:
        zf.writestr('[Content_Types].xml', ct)
        zf.writestr('_rels/.rels', rels)
        zf.writestr('xl/workbook.xml', wb)
        zf.writestr('xl/_rels/workbook.xml.rels', wb_rels)
        zf.writestr('xl/styles.xml', styles_xml)
        zf.writestr('xl/sharedStrings.xml', shared_strings_xml)
        zf.writestr('xl/theme/theme1.xml', theme_xml)
        zf.writestr('xl/worksheets/sheet1.xml', sheet1_xml)
        zf.writestr('xl/worksheets/sheet2.xml', sheet2_xml)
        zf.writestr('xl/worksheets/sheet3.xml', sheet3_xml)
        zf.writestr('xl/worksheets/sheet4.xml', sheet4_xml)
        zf.writestr('xl/worksheets/sheet5.xml', sheet5_xml)
        zf.writestr('xl/worksheets/sheet6.xml', sheet6_xml)
        zf.writestr('xl/worksheets/sheet7.xml', sheet7_xml)

    print(f'Output: {OUTPUT_XLSX}')

    # Verify
    with zipfile.ZipFile(OUTPUT_XLSX) as z:
        print(f'  Valid xlsx with {len(z.namelist())} files')
        print(f'  Contents: {z.namelist()}')

    # Verify formulas present
    import re
    with zipfile.ZipFile(OUTPUT_XLSX) as z:
        s1 = z.read('xl/worksheets/sheet1.xml').decode()
        formulas = re.findall(r'<f>(.*?)</f>', s1)
        print(f'  KALKULATOR sheet formulas: {len(formulas)}')
        if formulas:
            print(f'  Sample: {formulas[0][:80]}...')
        dv = 'dataValidation' in s1
        print(f'  Data validation present: {dv}')


if __name__ == "__main__":
    main()
