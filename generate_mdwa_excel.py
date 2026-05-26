#!/usr/bin/env python3
"""Generate Kalkulator_MDWA.xlsx - MDWA product presentation for insurance agents."""
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


def build_shared_strings(strings):
    ss = chr(60) + '?xml version="1.0" encoding="UTF-8" standalone="yes"?' + chr(62)
    ss += '<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"'
    ss += ' count="' + str(len(strings)) + '" uniqueCount="' + str(len(strings)) + '">'
    for s in strings:
        escaped = xml_escape(s)
        ss += '<si><t>' + escaped + '</t></si>'
    ss += '</sst>'
    return ss


def build_styles_xml():
    s = chr(60) + '?xml version="1.0" encoding="UTF-8" standalone="yes"?' + chr(62)
    s += '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
    # numFmts
    s += '<numFmts count="2">'
    s += '<numFmt numFmtId="164" formatCode="#,##0"/>'
    s += '<numFmt numFmtId="165" formatCode="0.00%"/>'
    s += '</numFmts>'
    # fonts: 0=default, 1=bold white, 2=bold green title, 3=bold black, 4=normal small, 5=bold white large
    s += '<fonts count="6">'
    s += '<font><sz val="10.0"/><color rgb="FF000000"/><name val="Arial"/></font>'
    s += '<font><b/><sz val="10.0"/><color rgb="FFFFFFFF"/><name val="Arial"/></font>'
    s += '<font><b/><sz val="14.0"/><color rgb="FF1E7145"/><name val="Arial"/></font>'
    s += '<font><b/><sz val="11.0"/><color rgb="FF000000"/><name val="Arial"/></font>'
    s += '<font><sz val="9.0"/><color rgb="FF333333"/><name val="Arial"/></font>'
    s += '<font><b/><sz val="12.0"/><color rgb="FFFFFFFF"/><name val="Arial"/></font>'
    s += '</fonts>'
    # fills: 0=none, 1=gray, 2=manulife green, 3=light green, 4=white, 5=dark green header
    s += '<fills count="6">'
    s += '<fill><patternFill patternType="none"/></fill>'
    s += '<fill><patternFill patternType="lightGray"/></fill>'
    s += '<fill><patternFill patternType="solid"><fgColor rgb="FF1E7145"/><bgColor rgb="FF1E7145"/></patternFill></fill>'
    s += '<fill><patternFill patternType="solid"><fgColor rgb="FFE8F5EE"/><bgColor rgb="FFE8F5EE"/></patternFill></fill>'
    s += '<fill><patternFill patternType="solid"><fgColor rgb="FFFFFFFF"/><bgColor rgb="FFFFFFFF"/></patternFill></fill>'
    s += '<fill><patternFill patternType="solid"><fgColor rgb="FF00B050"/><bgColor rgb="FF00B050"/></patternFill></fill>'
    s += '</fills>'
    # borders
    s += '<borders count="2">'
    s += '<border/>'
    s += '<border><left style="thin"><color rgb="FFCCCCCC"/></left>'
    s += '<right style="thin"><color rgb="FFCCCCCC"/></right>'
    s += '<top style="thin"><color rgb="FFCCCCCC"/></top>'
    s += '<bottom style="thin"><color rgb="FFCCCCCC"/></bottom></border>'
    s += '</borders>'
    # cellStyleXfs
    s += '<cellStyleXfs count="1">'
    s += '<xf borderId="0" fillId="0" fontId="0" numFmtId="0"/>'
    s += '</cellStyleXfs>'
    # cellXfs
    # 0: default
    # 1: header green bg white bold text centered
    # 2: data cell with border centered
    # 3: title bold green font
    # 4: data cell light green bg
    # 5: section header dark green bg white bold large
    # 6: number format with comma
    # 7: bold black left aligned
    s += '<cellXfs count="8">'
    s += '<xf borderId="0" fillId="0" fontId="0" numFmtId="0" xfId="0" applyAlignment="1"><alignment vertical="center" wrapText="1"/></xf>'
    s += '<xf borderId="1" fillId="2" fontId="1" numFmtId="0" xfId="0" applyAlignment="1" applyBorder="1" applyFill="1" applyFont="1"><alignment horizontal="center" vertical="center" wrapText="1"/></xf>'
    s += '<xf borderId="1" fillId="0" fontId="0" numFmtId="0" xfId="0" applyAlignment="1" applyBorder="1"><alignment horizontal="center" vertical="center" wrapText="1"/></xf>'
    s += '<xf borderId="0" fillId="0" fontId="2" numFmtId="0" xfId="0" applyAlignment="1" applyFont="1"><alignment vertical="center" wrapText="1"/></xf>'
    s += '<xf borderId="1" fillId="3" fontId="0" numFmtId="0" xfId="0" applyAlignment="1" applyBorder="1" applyFill="1"><alignment horizontal="center" vertical="center" wrapText="1"/></xf>'
    s += '<xf borderId="1" fillId="5" fontId="5" numFmtId="0" xfId="0" applyAlignment="1" applyBorder="1" applyFill="1" applyFont="1"><alignment horizontal="center" vertical="center" wrapText="1"/></xf>'
    s += '<xf borderId="1" fillId="0" fontId="0" numFmtId="164" xfId="0" applyAlignment="1" applyBorder="1" applyNumberFormat="1"><alignment horizontal="center" vertical="center" wrapText="1"/></xf>'
    s += '<xf borderId="0" fillId="0" fontId="3" numFmtId="0" xfId="0" applyAlignment="1" applyFont="1"><alignment vertical="center" wrapText="1"/></xf>'
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


class SheetBuilder:
    """Helper to build worksheet XML with rows and cells."""

    def __init__(self, ss_map):
        self.ss_map = ss_map
        self.rows = []
        self.merges = []
        self.col_widths = {}

    def set_col_width(self, col, width):
        self.col_widths[col] = width

    def add_row(self, row_num, cells, height=None):
        self.rows.append((row_num, cells, height))

    def add_merge(self, ref):
        self.merges.append(ref)

    def _cell_xml(self, col_idx, row_num, value, style=0):
        ref = col_letter(col_idx) + str(row_num)
        if isinstance(value, str):
            if value in self.ss_map:
                si = self.ss_map[value]
                return '<c r="' + ref + '" s="' + str(style) + '" t="s"><v>' + str(si) + '</v></c>'
            else:
                # inline string
                return '<c r="' + ref + '" s="' + str(style) + '" t="inlineStr"><is><t>' + xml_escape(value) + '</t></is></c>'
        elif isinstance(value, (int, float)):
            return '<c r="' + ref + '" s="' + str(style) + '"><v>' + str(value) + '</v></c>'
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
            for col_idx, value, style in cells:
                s += self._cell_xml(col_idx, row_num, value, style)
            s += '</row>'
        s += '</sheetData>'
        if self.merges:
            s += '<mergeCells count="' + str(len(self.merges)) + '">'
            for m in self.merges:
                s += '<mergeCell ref="' + m + '"/>'
            s += '</mergeCells>'
        s += '</worksheet>'
        return s


def build_overview_sheet(ss_map):
    """Sheet 1: Overview - Product Summary comparing all 3 MDWA plans."""
    sb = SheetBuilder(ss_map)
    sb.set_col_width(1, 25)
    sb.set_col_width(2, 30)
    sb.set_col_width(3, 30)
    sb.set_col_width(4, 30)

    r = 1
    sb.add_row(r, [(0, 'MANULIFE DYNAMIC WEALTH ASSURANCE (MDWA)', 5)], 30)
    sb.add_merge('A1:D1')
    r += 1
    sb.add_row(r, [(0, 'Perbandingan 3 Plan MDWA', 3)], 24)
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
    sb.add_row(r, [(0, 'SPESIFIKASI UMUM', 5)], 24)
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
    sb.add_row(r, [(0, 'PREMIUM BANDING', 5)], 24)
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


def build_plan_a_sheet(ss_map):
    """Sheet 2: Plan A - Anuitas - Annual Payout benefits."""
    sb = SheetBuilder(ss_map)
    sb.set_col_width(1, 22)
    sb.set_col_width(2, 22)
    sb.set_col_width(3, 22)

    r = 1
    sb.add_row(r, [(0, 'PLAN A - ANUITAS (Annual Payout)', 5)], 28)
    sb.add_merge('A1:C1')
    r += 2

    # Banding 1 - Coverage 20 years
    sb.add_row(r, [(0, 'BANDING 1 - Masa Pertanggungan 20 Tahun', 7)])
    r += 1
    sb.add_row(r, [(0, 'Masa Pembayaran', 1), (1, 'Annual Payout (IDR)', 1), (2, 'Annual Payout (USD)', 1)])
    r += 1
    sb.add_row(r, [(0, 'Premi Sekaligus', 2), (1, '9.5% of Single Premium', 2), (2, '8.25% of Single Premium', 2)])
    r += 1
    sb.add_row(r, [(0, '5 tahun', 2), (1, '65% of Annual Premium', 2), (2, '55% of Annual Premium', 2)])
    r += 1
    sb.add_row(r, [(0, '10 tahun', 2), (1, '235% of Annual Premium', 2), (2, '205% of Annual Premium', 2)])
    r += 2

    # Banding 1 - Coverage 30 years
    sb.add_row(r, [(0, 'BANDING 1 - Masa Pertanggungan 30 Tahun', 7)])
    r += 1
    sb.add_row(r, [(0, 'Masa Pembayaran', 1), (1, 'Annual Payout (IDR)', 1), (2, 'Annual Payout (USD)', 1)])
    r += 1
    sb.add_row(r, [(0, 'Premi Sekaligus', 2), (1, '7% of Single Premium', 2), (2, '5.5% of Single Premium', 2)])
    r += 1
    sb.add_row(r, [(0, '5 tahun', 2), (1, '40% of Annual Premium', 2), (2, '32.5% of Annual Premium', 2)])
    r += 1
    sb.add_row(r, [(0, '10 tahun', 2), (1, '105% of Annual Premium', 2), (2, '85% of Annual Premium', 2)])
    r += 2

    # Banding 2 Extra - Coverage 20 years
    sb.add_row(r, [(0, 'BANDING 2 EXTRA - Masa Pertanggungan 20 Tahun', 7)])
    r += 1
    sb.add_row(r, [(0, 'Masa Pembayaran', 1), (1, 'Extra (IDR)', 1), (2, 'Extra (USD)', 1)])
    r += 1
    sb.add_row(r, [(0, 'Premi Sekaligus', 2), (1, '+6%', 2), (2, '+3%', 2)])
    r += 1
    sb.add_row(r, [(0, '5 tahun', 2), (1, '+30%', 2), (2, '+15%', 2)])
    r += 1
    sb.add_row(r, [(0, '10 tahun', 2), (1, '+60%', 2), (2, '+30%', 2)])
    r += 2

    # Banding 2 Extra - Coverage 30 years
    sb.add_row(r, [(0, 'BANDING 2 EXTRA - Masa Pertanggungan 30 Tahun', 7)])
    r += 1
    sb.add_row(r, [(0, 'Masa Pembayaran', 1), (1, 'Extra (IDR)', 1), (2, 'Extra (USD)', 1)])
    r += 1
    sb.add_row(r, [(0, 'Premi Sekaligus', 2), (1, '+9%', 2), (2, '+4.5%', 2)])
    r += 1
    sb.add_row(r, [(0, '5 tahun', 2), (1, '+45%', 2), (2, '+22.5%', 2)])
    r += 1
    sb.add_row(r, [(0, '10 tahun', 2), (1, '+90%', 2), (2, '+45%', 2)])
    r += 2

    # Payout Schedule
    sb.add_row(r, [(0, 'JADWAL PEMBAYARAN TAHAPAN', 7)])
    r += 1
    sb.add_row(r, [(0, 'Masa Pembayaran', 1), (1, 'Mulai Pembayaran', 1), (2, 'Sampai', 1)])
    r += 1
    sb.add_row(r, [(0, 'Premi Sekaligus', 2), (1, 'Akhir tahun polis ke-6', 2), (2, 'Akhir masa pertanggungan', 2)])
    r += 1
    sb.add_row(r, [(0, '5 tahun', 2), (1, 'Akhir tahun polis ke-10', 2), (2, 'Akhir masa pertanggungan', 2)])
    r += 1
    sb.add_row(r, [(0, '10 tahun', 2), (1, 'Akhir tahun polis ke-15', 2), (2, 'Akhir masa pertanggungan', 2)])

    return sb.build()


def build_plan_b_sheet(ss_map):
    """Sheet 3: Plan B - Beasiswa - Lump Sum Maturity."""
    sb = SheetBuilder(ss_map)
    sb.set_col_width(1, 18)
    sb.set_col_width(2, 16)
    sb.set_col_width(3, 16)

    r = 1
    sb.add_row(r, [(0, 'PLAN B - BEASISWA (Lump Sum at Maturity)', 5)], 28)
    sb.add_merge('A1:C1')
    r += 2

    # Banding 1 - Single Premium
    sb.add_row(r, [(0, 'BANDING 1 - Premi Sekaligus (% of Single Premium)', 7)])
    r += 1
    sb.add_row(r, [(0, 'Coverage (Tahun)', 1), (1, 'IDR', 1), (2, 'USD', 1)])
    r += 1
    single_data = [
        (5, '110%', '105%'), (6, '115%', '106%'), (7, '120%', '108%'),
        (8, '125%', '110%'), (9, '130%', '112%'), (10, '135%', '114%'),
        (11, '140%', '116%'), (12, '145%', '120%'), (13, '150%', '125%'),
        (14, '155%', '130%'), (15, '160%', '135%'), (20, '200%', '160%'),
        (30, '325%', '200%'),
    ]
    for cov, idr, usd in single_data:
        sb.add_row(r, [(0, str(cov), 2), (1, idr, 2), (2, usd, 2)])
        r += 1
    r += 1

    # Banding 1 - PPP 5 tahun
    sb.add_row(r, [(0, 'BANDING 1 - PPP 5 Tahun (% of Annual Premium)', 7)])
    r += 1
    sb.add_row(r, [(0, 'Coverage (Tahun)', 1), (1, 'IDR', 1), (2, 'USD', 1)])
    r += 1
    ppp5_data = [
        (10, '600%', '535%'), (11, '625%', '545%'), (12, '650%', '555%'),
        (13, '675%', '565%'), (14, '700%', '575%'), (15, '725%', '600%'),
        (20, '950%', '700%'), (30, '1625%', '950%'),
    ]
    for cov, idr, usd in ppp5_data:
        sb.add_row(r, [(0, str(cov), 2), (1, idr, 2), (2, usd, 2)])
        r += 1
    r += 1

    # Banding 1 - PPP 10 tahun
    sb.add_row(r, [(0, 'BANDING 1 - PPP 10 Tahun (% of Annual Premium)', 7)])
    r += 1
    sb.add_row(r, [(0, 'Coverage (Tahun)', 1), (1, 'IDR', 1), (2, 'USD', 1)])
    r += 1
    ppp10_data = [
        (15, '1350%', '1175%'), (20, '1800%', '1350%'), (30, '3250%', '1800%'),
    ]
    for cov, idr, usd in ppp10_data:
        sb.add_row(r, [(0, str(cov), 2), (1, idr, 2), (2, usd, 2)])
        r += 1
    r += 2

    # Banding 2 Extra
    sb.add_row(r, [(0, 'BANDING 2 EXTRA (tambahan di atas Banding 1)', 5)], 24)
    sb.add_merge('A' + str(r) + ':F' + str(r))
    r += 1
    sb.set_col_width(4, 16)
    sb.set_col_width(5, 16)
    sb.set_col_width(6, 16)
    sb.set_col_width(7, 16)
    sb.add_row(r, [(0, 'Coverage', 1), (1, 'Single IDR', 1), (2, 'Single USD', 1),
                   (3, '5yr IDR', 1), (4, '5yr USD', 1), (5, '10yr IDR', 1), (6, '10yr USD', 1)])
    r += 1
    banding2_data = [
        (5, '+1.50%', '+0.75%', '-', '-', '-', '-'),
        (6, '+1.80%', '+0.90%', '-', '-', '-', '-'),
        (7, '+2.10%', '+1.05%', '-', '-', '-', '-'),
        (8, '+2.40%', '+1.20%', '-', '-', '-', '-'),
        (9, '+2.70%', '+1.35%', '-', '-', '-', '-'),
        (10, '+3.00%', '+1.50%', '+15.00%', '+7.50%', '-', '-'),
        (11, '+3.30%', '+1.65%', '+16.50%', '+8.25%', '-', '-'),
        (12, '+3.60%', '+1.80%', '+18.00%', '+9.00%', '-', '-'),
        (13, '+3.90%', '+1.95%', '+19.50%', '+9.75%', '-', '-'),
        (14, '+4.20%', '+2.10%', '+21.00%', '+10.50%', '-', '-'),
        (15, '+4.50%', '+2.25%', '+22.50%', '+11.25%', '+45.00%', '+22.50%'),
        (20, '+6.00%', '+3.00%', '+30.00%', '+15.00%', '+60.00%', '+30.00%'),
        (30, '+9.00%', '+4.50%', '+45.00%', '+22.50%', '+90.00%', '+45.00%'),
    ]
    for row_data in banding2_data:
        cov = row_data[0]
        sb.add_row(r, [(0, str(cov), 2), (1, row_data[1], 2), (2, row_data[2], 2),
                       (3, row_data[3], 2), (4, row_data[4], 2), (5, row_data[5], 2), (6, row_data[6], 2)])
        r += 1

    return sb.build()


def build_plan_c_sheet(ss_map):
    """Sheet 4: Plan C - Combo - Annual Payout + Maturity."""
    sb = SheetBuilder(ss_map)
    for i in range(1, 10):
        sb.set_col_width(i, 16)

    r = 1
    sb.add_row(r, [(0, 'PLAN C - COMBO (Annual Payout + Maturity)', 5)], 28)
    sb.add_merge('A1:H1')
    r += 2

    # Banding 1 IDR
    sb.add_row(r, [(0, 'BANDING 1 - IDR', 7)])
    r += 1
    sb.add_row(r, [(0, 'Pembayaran', 1), (1, '10yr Tahapan', 1), (2, '10yr Maturity', 1),
                   (3, '15yr Tahapan', 1), (4, '15yr Maturity', 1),
                   (5, '20yr Tahapan', 1), (6, '20yr Maturity', 1),
                   (7, '30yr Tahapan', 1), (8, '30yr Maturity', 1)])
    r += 1
    sb.add_row(r, [(0, 'Single', 2), (1, '2%', 2), (2, '105%', 2),
                   (3, '2%', 2), (4, '120%', 2), (5, '2%', 2), (6, '135%', 2),
                   (7, '2%', 2), (8, '170%', 2)])
    r += 1
    sb.add_row(r, [(0, '5 tahun', 2), (1, '13.5%', 2), (2, '510%', 2),
                   (3, '17.5%', 2), (4, '510%', 2), (5, '20%', 2), (6, '510%', 2),
                   (7, '20%', 2), (8, '510%', 2)])
    r += 1
    sb.add_row(r, [(0, '10 tahun', 2), (1, 'N/A', 2), (2, 'N/A', 2),
                   (3, '40%', 2), (4, '1000%', 2), (5, '40%', 2), (6, '1100%', 2),
                   (7, '40%', 2), (8, '1200%', 2)])
    r += 2

    # Banding 1 USD
    sb.add_row(r, [(0, 'BANDING 1 - USD', 7)])
    r += 1
    sb.add_row(r, [(0, 'Pembayaran', 1), (1, '10yr Tahapan', 1), (2, '10yr Maturity', 1),
                   (3, '15yr Tahapan', 1), (4, '15yr Maturity', 1),
                   (5, '20yr Tahapan', 1), (6, '20yr Maturity', 1),
                   (7, '30yr Tahapan', 1), (8, '30yr Maturity', 1)])
    r += 1
    sb.add_row(r, [(0, 'Single', 2), (1, '1%', 2), (2, '100%', 2),
                   (3, '1%', 2), (4, '110%', 2), (5, '1%', 2), (6, '120%', 2),
                   (7, '1%', 2), (8, '140%', 2)])
    r += 1
    sb.add_row(r, [(0, '5 tahun', 2), (1, '5%', 2), (2, '500%', 2),
                   (3, '8%', 2), (4, '500%', 2), (5, '11%', 2), (6, '500%', 2),
                   (7, '11%', 2), (8, '500%', 2)])
    r += 1
    sb.add_row(r, [(0, '10 tahun', 2), (1, 'N/A', 2), (2, 'N/A', 2),
                   (3, '20%', 2), (4, '1000%', 2), (5, '20%', 2), (6, '1100%', 2),
                   (7, '20%', 2), (8, '1150%', 2)])
    r += 2

    # Banding 2 Extra IDR (Maturity only)
    sb.add_row(r, [(0, 'BANDING 2 EXTRA - Maturity IDR', 7)])
    r += 1
    sb.add_row(r, [(0, 'Pembayaran', 1), (1, '10yr', 1), (2, '15yr', 1), (3, '20yr', 1), (4, '30yr', 1)])
    r += 1
    sb.add_row(r, [(0, 'Single', 2), (1, '+3%', 2), (2, '+5%', 2), (3, '+6%', 2), (4, '+9%', 2)])
    r += 1
    sb.add_row(r, [(0, '5 tahun', 2), (1, '+15%', 2), (2, '+23%', 2), (3, '+30%', 2), (4, '+45%', 2)])
    r += 1
    sb.add_row(r, [(0, '10 tahun', 2), (1, 'N/A', 2), (2, '+45%', 2), (3, '+60%', 2), (4, '+90%', 2)])
    r += 2

    # Banding 2 Extra USD
    sb.add_row(r, [(0, 'BANDING 2 EXTRA - Maturity USD', 7)])
    r += 1
    sb.add_row(r, [(0, 'Pembayaran', 1), (1, '10yr', 1), (2, '15yr', 1), (3, '20yr', 1), (4, '30yr', 1)])
    r += 1
    sb.add_row(r, [(0, 'Single', 2), (1, '+1%', 2), (2, '+2.25%', 2), (3, '+3%', 2), (4, '+5%', 2)])
    r += 1
    sb.add_row(r, [(0, '5 tahun', 2), (1, '+8%', 2), (2, '+11.25%', 2), (3, '+15%', 2), (4, '+23%', 2)])
    r += 1
    sb.add_row(r, [(0, '10 tahun', 2), (1, 'N/A', 2), (2, '+23%', 2), (3, '+30%', 2), (4, '+45%', 2)])

    return sb.build()


def build_simulasi_sheet(ss_map):
    """Sheet 5: Simulasi - Simulation/Calculator for agents."""
    sb = SheetBuilder(ss_map)
    sb.set_col_width(1, 25)
    sb.set_col_width(2, 20)
    sb.set_col_width(3, 20)
    sb.set_col_width(4, 20)
    sb.set_col_width(5, 20)

    r = 1
    sb.add_row(r, [(0, 'SIMULASI MANFAAT MDWA', 5)], 28)
    sb.add_merge('A1:E1')
    r += 2

    # Simulation 1: Single Premium IDR 500M (Banding 2)
    sb.add_row(r, [(0, 'SIMULASI 1: Premi Sekaligus IDR 500.000.000 (Banding 2)', 7)])
    r += 2
    sb.add_row(r, [(0, 'Plan', 1), (1, 'Coverage', 1), (2, 'Manfaat Tahunan', 1), (3, 'Manfaat Jatuh Tempo', 1), (4, 'Total Manfaat', 1)])
    r += 1
    # Plan A - Single 20yr: 9.5%+6% = 15.5% of 500M = 77.5M/yr x 15 years (yr 6-20)
    sb.add_row(r, [(0, 'Plan A - Anuitas', 4), (1, '20 tahun', 2), (2, '77.5M/tahun x 15 thn = 1.163B', 2), (3, '-', 2), (4, 'IDR 1.162.500.000', 2)])
    r += 1
    # Plan A - Single 30yr: 7%+9% = 16% of 500M = 80M/yr x 25 years (yr 6-30)
    sb.add_row(r, [(0, 'Plan A - Anuitas', 4), (1, '30 tahun', 2), (2, '80M/tahun x 25 thn = 2B', 2), (3, '-', 2), (4, 'IDR 2.000.000.000', 2)])
    r += 1
    # Plan B - Single 20yr: 200%+6% = 206% of 500M = 1.03B
    sb.add_row(r, [(0, 'Plan B - Beasiswa', 4), (1, '20 tahun', 2), (2, '-', 2), (3, 'IDR 1.030.000.000', 2), (4, 'IDR 1.030.000.000', 2)])
    r += 1
    # Plan B - Single 30yr: 325%+9% = 334% of 500M = 1.67B
    sb.add_row(r, [(0, 'Plan B - Beasiswa', 4), (1, '30 tahun', 2), (2, '-', 2), (3, 'IDR 1.670.000.000', 2), (4, 'IDR 1.670.000.000', 2)])
    r += 1
    # Plan C - Single 20yr: Tahapan 2% x 500M = 10M/yr (yr6-20=15yrs=150M) + Maturity (135%+6%)=141%=705M
    sb.add_row(r, [(0, 'Plan C - Combo', 4), (1, '20 tahun', 2), (2, '10M/tahun x 15 thn = 150M', 2), (3, 'IDR 705.000.000', 2), (4, 'IDR 855.000.000', 2)])
    r += 1
    # Plan C - Single 30yr: Tahapan 2% x 500M = 10M/yr (yr6-30=25yrs=250M) + Maturity (170%+9%)=179%=895M
    sb.add_row(r, [(0, 'Plan C - Combo', 4), (1, '30 tahun', 2), (2, '10M/tahun x 25 thn = 250M', 2), (3, 'IDR 895.000.000', 2), (4, 'IDR 1.145.000.000', 2)])
    r += 2

    # Simulation 2: Regular Premium IDR 100M/yr, PPP 5 tahun (Banding 2)
    sb.add_row(r, [(0, 'SIMULASI 2: Premi Tahunan IDR 100.000.000, PPP 5 Tahun (Banding 2)', 7)])
    r += 2
    sb.add_row(r, [(0, 'Plan', 1), (1, 'Coverage', 1), (2, 'Manfaat Tahunan', 1), (3, 'Manfaat Jatuh Tempo', 1), (4, 'Total Manfaat', 1)])
    r += 1
    # Plan A - 5yr 20yr: 65%+30% = 95% of 100M = 95M/yr x 11 years (yr 10-20)
    sb.add_row(r, [(0, 'Plan A - Anuitas', 4), (1, '20 tahun', 2), (2, '95M/tahun x 11 thn = 1.045B', 2), (3, '-', 2), (4, 'IDR 1.045.000.000', 2)])
    r += 1
    # Plan A - 5yr 30yr: 40%+45% = 85% of 100M = 85M/yr x 21 years (yr 10-30)
    sb.add_row(r, [(0, 'Plan A - Anuitas', 4), (1, '30 tahun', 2), (2, '85M/tahun x 21 thn = 1.785B', 2), (3, '-', 2), (4, 'IDR 1.785.000.000', 2)])
    r += 1
    # Plan B - 5yr 20yr: 950%+30% = 980% of 100M = 980M
    sb.add_row(r, [(0, 'Plan B - Beasiswa', 4), (1, '20 tahun', 2), (2, '-', 2), (3, 'IDR 980.000.000', 2), (4, 'IDR 980.000.000', 2)])
    r += 1
    # Plan B - 5yr 30yr: 1625%+45% = 1670% of 100M = 1.67B
    sb.add_row(r, [(0, 'Plan B - Beasiswa', 4), (1, '30 tahun', 2), (2, '-', 2), (3, 'IDR 1.670.000.000', 2), (4, 'IDR 1.670.000.000', 2)])
    r += 1
    # Plan C - 5yr 20yr: Tahapan 20% of 100M = 20M/yr x 11yrs (yr10-20)=220M + Maturity (510%+30%)=540%=540M
    sb.add_row(r, [(0, 'Plan C - Combo', 4), (1, '20 tahun', 2), (2, '20M/tahun x 11 thn = 220M', 2), (3, 'IDR 540.000.000', 2), (4, 'IDR 760.000.000', 2)])
    r += 1
    # Plan C - 5yr 30yr: Tahapan 20% of 100M = 20M/yr x 21yrs (yr10-30)=420M + Maturity (510%+45%)=555%=555M
    sb.add_row(r, [(0, 'Plan C - Combo', 4), (1, '30 tahun', 2), (2, '20M/tahun x 21 thn = 420M', 2), (3, 'IDR 555.000.000', 2), (4, 'IDR 975.000.000', 2)])
    r += 2

    # Simulation 3: Regular Premium IDR 100M/yr, PPP 10 tahun (Banding 1)
    sb.add_row(r, [(0, 'SIMULASI 3: Premi Tahunan IDR 100.000.000, PPP 10 Tahun (Banding 1)', 7)])
    r += 2
    sb.add_row(r, [(0, 'Plan', 1), (1, 'Coverage', 1), (2, 'Manfaat Tahunan', 1), (3, 'Manfaat Jatuh Tempo', 1), (4, 'Total Manfaat', 1)])
    r += 1
    # Plan A - 10yr 20yr: 235% of 100M = 235M/yr x 6 years (yr 15-20)
    sb.add_row(r, [(0, 'Plan A - Anuitas', 4), (1, '20 tahun', 2), (2, '235M/tahun x 6 thn = 1.41B', 2), (3, '-', 2), (4, 'IDR 1.410.000.000', 2)])
    r += 1
    # Plan A - 10yr 30yr: 105% of 100M = 105M/yr x 16 years (yr 15-30)
    sb.add_row(r, [(0, 'Plan A - Anuitas', 4), (1, '30 tahun', 2), (2, '105M/tahun x 16 thn = 1.68B', 2), (3, '-', 2), (4, 'IDR 1.680.000.000', 2)])
    r += 1
    # Plan B - 10yr 20yr: 1800% of 100M = 1.8B
    sb.add_row(r, [(0, 'Plan B - Beasiswa', 4), (1, '20 tahun', 2), (2, '-', 2), (3, 'IDR 1.800.000.000', 2), (4, 'IDR 1.800.000.000', 2)])
    r += 1
    # Plan B - 10yr 30yr: 3250% of 100M = 3.25B
    sb.add_row(r, [(0, 'Plan B - Beasiswa', 4), (1, '30 tahun', 2), (2, '-', 2), (3, 'IDR 3.250.000.000', 2), (4, 'IDR 3.250.000.000', 2)])
    r += 1
    # Plan C - 10yr 20yr: Tahapan 40% of 100M = 40M/yr x 6yrs (yr15-20)=240M + Maturity 1100%=1.1B
    sb.add_row(r, [(0, 'Plan C - Combo', 4), (1, '20 tahun', 2), (2, '40M/tahun x 6 thn = 240M', 2), (3, 'IDR 1.100.000.000', 2), (4, 'IDR 1.340.000.000', 2)])
    r += 1
    # Plan C - 10yr 30yr: Tahapan 40% of 100M = 40M/yr x 16yrs (yr15-30)=640M + Maturity 1200%=1.2B
    sb.add_row(r, [(0, 'Plan C - Combo', 4), (1, '30 tahun', 2), (2, '40M/tahun x 16 thn = 640M', 2), (3, 'IDR 1.200.000.000', 2), (4, 'IDR 1.840.000.000', 2)])

    return sb.build()


def build_komisi_sheet(ss_map):
    """Sheet 6: Komisi - Commission Structure."""
    sb = SheetBuilder(ss_map)
    sb.set_col_width(1, 20)
    sb.set_col_width(2, 18)
    sb.set_col_width(3, 15)
    sb.set_col_width(4, 15)

    r = 1
    sb.add_row(r, [(0, 'STRUKTUR KOMISI MDWA', 5)], 28)
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
        ('Single', '20yr', '4%', '-'),
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
    """Sheet 7: Perbandingan - Comparison vs Competitors."""
    sb = SheetBuilder(ss_map)
    for i in range(1, 8):
        sb.set_col_width(i, 18)

    r = 1
    sb.add_row(r, [(0, 'PERBANDINGAN vs KOMPETITOR', 5)], 28)
    sb.add_merge('A1:G1')
    r += 2

    sb.add_row(r, [(0, 'PPP 5 Tahun, Band 2, Premi Tahunan IDR 100 Juta', 7)])
    r += 2

    # Header
    sb.add_row(r, [(0, 'Feature', 1), (1, 'MDWA 15yr', 1), (2, 'MDWA 20yr', 1),
                   (3, 'MSP 15yr', 1), (4, 'MSP 20yr', 1), (5, 'FI 15yr', 1), (6, 'FI 20yr', 1)])
    r += 1
    # Data rows
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


def main():
    print('Generating Kalkulator_MDWA.xlsx...')

    # Collect all unique strings used across all sheets
    all_strings = []
    seen = set()

    def add_str(s):
        if s not in seen:
            seen.add(s)
            all_strings.append(s)

    # Overview strings
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

    # Plan A strings
    plan_a_strs = [
        'PLAN A - ANUITAS (Annual Payout)',
        'BANDING 1 - Masa Pertanggungan 20 Tahun',
        'Masa Pembayaran', 'Annual Payout (IDR)', 'Annual Payout (USD)',
        'Premi Sekaligus', '9.5% of Single Premium', '8.25% of Single Premium',
        '5 tahun', '65% of Annual Premium', '55% of Annual Premium',
        '10 tahun', '235% of Annual Premium', '205% of Annual Premium',
        'BANDING 1 - Masa Pertanggungan 30 Tahun',
        '7% of Single Premium', '5.5% of Single Premium',
        '40% of Annual Premium', '32.5% of Annual Premium',
        '105% of Annual Premium', '85% of Annual Premium',
        'BANDING 2 EXTRA - Masa Pertanggungan 20 Tahun',
        'Extra (IDR)', 'Extra (USD)',
        '+6%', '+3%', '+30%', '+15%', '+60%',
        'BANDING 2 EXTRA - Masa Pertanggungan 30 Tahun',
        '+9%', '+4.5%', '+45%', '+22.5%', '+90%',
        'JADWAL PEMBAYARAN TAHAPAN',
        'Mulai Pembayaran', 'Sampai',
        'Akhir tahun polis ke-6', 'Akhir masa pertanggungan',
        'Akhir tahun polis ke-10', 'Akhir tahun polis ke-15',
    ]
    for s in plan_a_strs:
        add_str(s)

    # Plan B strings
    plan_b_strs = [
        'PLAN B - BEASISWA (Lump Sum at Maturity)',
        'BANDING 1 - Premi Sekaligus (% of Single Premium)',
        'Coverage (Tahun)', 'IDR', 'USD',
        '5', '110%', '105%', '6', '115%', '106%', '7', '120%', '108%',
        '8', '125%', '110%', '9', '130%', '112%', '10', '135%', '114%',
        '11', '140%', '116%', '12', '145%', '120%', '13', '150%', '125%',
        '14', '155%', '130%', '15', '160%', '135%', '20', '200%', '160%',
        '30', '325%', '200%',
        'BANDING 1 - PPP 5 Tahun (% of Annual Premium)',
        '600%', '535%', '625%', '545%', '650%', '555%',
        '675%', '565%', '700%', '575%', '725%', '600%',
        '950%', '700%', '1625%', '950%',
        'BANDING 1 - PPP 10 Tahun (% of Annual Premium)',
        '1350%', '1175%', '1800%', '1350%', '3250%',
        'BANDING 2 EXTRA (tambahan di atas Banding 1)',
        'Coverage', 'Single IDR', 'Single USD', '5yr IDR', '5yr USD', '10yr IDR', '10yr USD',
        '+1.50%', '+0.75%', '-',
        '+1.80%', '+0.90%',
        '+2.10%', '+1.05%',
        '+2.40%', '+1.20%',
        '+2.70%', '+1.35%',
        '+3.00%', '+1.50%', '+15.00%', '+7.50%',
        '+3.30%', '+1.65%', '+16.50%', '+8.25%',
        '+3.60%', '+1.80%', '+18.00%', '+9.00%',
        '+3.90%', '+1.95%', '+19.50%', '+9.75%',
        '+4.20%', '+2.10%', '+21.00%', '+10.50%',
        '+4.50%', '+2.25%', '+22.50%', '+11.25%', '+45.00%',
        '+6.00%', '+3.00%', '+30.00%', '+15.00%', '+60.00%', '+30.00%',
        '+9.00%', '+45.00%',
    ]
    for s in plan_b_strs:
        add_str(s)

    # Plan C strings
    plan_c_strs = [
        'PLAN C - COMBO (Annual Payout + Maturity)',
        'BANDING 1 - IDR', 'Pembayaran',
        '10yr Tahapan', '10yr Maturity', '15yr Tahapan', '15yr Maturity',
        '20yr Tahapan', '20yr Maturity', '30yr Tahapan', '30yr Maturity',
        'Single', '2%', '105%', '120%', '135%', '170%',
        '13.5%', '510%', '17.5%', '20%',
        'N/A', '40%', '1000%', '1100%', '1200%',
        'BANDING 1 - USD',
        '1%', '100%', '110%', '140%',
        '5%', '500%', '8%', '11%',
        '1150%',
        'BANDING 2 EXTRA - Maturity IDR',
        '10yr', '15yr', '20yr', '30yr',
        '+5%', '+23%',
        'BANDING 2 EXTRA - Maturity USD',
        '+1%', '+2.25%', '+8%', '+11.25%',
    ]
    for s in plan_c_strs:
        add_str(s)

    # Simulasi strings
    sim_strs = [
        'SIMULASI MANFAAT MDWA',
        'SIMULASI 1: Premi Sekaligus IDR 500.000.000 (Banding 2)',
        'Plan', 'Manfaat Tahunan', 'Manfaat Jatuh Tempo', 'Total Manfaat',
        '20 tahun', '30 tahun',
        '77.5M/tahun x 15 thn = 1.163B', 'IDR 1.162.500.000',
        '80M/tahun x 25 thn = 2B', 'IDR 2.000.000.000',
        'IDR 1.030.000.000',
        'IDR 1.670.000.000',
        '10M/tahun x 15 thn = 150M', 'IDR 705.000.000', 'IDR 855.000.000',
        '10M/tahun x 25 thn = 250M', 'IDR 895.000.000', 'IDR 1.145.000.000',
        'SIMULASI 2: Premi Tahunan IDR 100.000.000, PPP 5 Tahun (Banding 2)',
        '95M/tahun x 11 thn = 1.045B', 'IDR 1.045.000.000',
        '85M/tahun x 21 thn = 1.785B', 'IDR 1.785.000.000',
        'IDR 980.000.000',
        '20M/tahun x 11 thn = 220M', 'IDR 540.000.000', 'IDR 760.000.000',
        '20M/tahun x 21 thn = 420M', 'IDR 555.000.000', 'IDR 975.000.000',
        'SIMULASI 3: Premi Tahunan IDR 100.000.000, PPP 10 Tahun (Banding 1)',
        '235M/tahun x 6 thn = 1.41B', 'IDR 1.410.000.000',
        '105M/tahun x 16 thn = 1.68B', 'IDR 1.680.000.000',
        'IDR 1.800.000.000',
        'IDR 3.250.000.000',
        '40M/tahun x 6 thn = 240M', 'IDR 1.100.000.000', 'IDR 1.340.000.000',
        '40M/tahun x 16 thn = 640M', 'IDR 1.200.000.000', 'IDR 1.840.000.000',
    ]
    for s in sim_strs:
        add_str(s)

    # Komisi strings
    komisi_strs = [
        'STRUKTUR KOMISI MDWA',
        'PLAN A - ANUITAS', 'Term', 'Year 1', 'Year 2',
        '4%', '5%', '10%', '15%', '12.5%', '7.5%',
        'PLAN B - BEASISWA',
        '5-10yr', '11-15yr', '2%', '6%', 'all',
        'PLAN C - COMBO',
    ]
    for s in komisi_strs:
        add_str(s)

    # Perbandingan strings
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

    ss_map = {s: i for i, s in enumerate(all_strings)}

    print('  Building sheets...')
    sheet1_xml = build_overview_sheet(ss_map)
    sheet2_xml = build_plan_a_sheet(ss_map)
    sheet3_xml = build_plan_b_sheet(ss_map)
    sheet4_xml = build_plan_c_sheet(ss_map)
    sheet5_xml = build_simulasi_sheet(ss_map)
    sheet6_xml = build_komisi_sheet(ss_map)
    sheet7_xml = build_perbandingan_sheet(ss_map)

    shared_strings_xml = build_shared_strings(all_strings)
    styles_xml = build_styles_xml()
    theme_xml = build_theme_xml()

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
    wb += '<sheet state="visible" name="Overview" sheetId="1" r:id="rId4"/>'
    wb += '<sheet state="visible" name="Plan A - Anuitas" sheetId="2" r:id="rId5"/>'
    wb += '<sheet state="visible" name="Plan B - Beasiswa" sheetId="3" r:id="rId6"/>'
    wb += '<sheet state="visible" name="Plan C - Combo" sheetId="4" r:id="rId7"/>'
    wb += '<sheet state="visible" name="Simulasi" sheetId="5" r:id="rId8"/>'
    wb += '<sheet state="visible" name="Komisi" sheetId="6" r:id="rId9"/>'
    wb += '<sheet state="visible" name="Perbandingan" sheetId="7" r:id="rId10"/>'
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


if __name__ == "__main__":
    main()
