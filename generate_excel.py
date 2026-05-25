#!/usr/bin/env python3
"""Generate Kalkulator_Premi_MiUHC.xlsx with 4 sheets from PDF premium data."""
import zipfile
import zlib
import re
import os

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
OUTPUT_XLSX = os.path.join(SCRIPT_DIR, 'Kalkulator_Premi_MiUHC.xlsx')
WA_LINK = "https://wa.me/6287781896087"

SYARIAH_PDF = os.path.join(SCRIPT_DIR,
    'B. Tabel Premi Syariah 2024 - Miuhc Versi Quick PDF (3).pdf')
KONV_PDF = os.path.join(SCRIPT_DIR,
    'B. Tabel Premi Miuhc Versi Quick PDF Per Sep 2025 (3).pdf')

KONV_INDICES = [0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 13]


def extract_numbers_from_stream(data, stream_idx):
    stream_pattern = rb'stream\r?\n(.*?)\r?\nendstream'
    matches = re.findall(stream_pattern, data, re.DOTALL)
    decompressed = zlib.decompress(matches[stream_idx]).decode('latin-1')
    tj_matches = re.findall(r'\[(.*?)\]\s*TJ', decompressed)
    text_parts = []
    for tj in tj_matches:
        parts = re.findall(r'\((.*?)\)', tj)
        text_parts.extend(parts)
    full_text = ''.join(text_parts)
    numbers = re.findall(r'\d{1,3}(?:,\d{3})+', full_text)
    return [int(n.replace(',', '')) for n in numbers]


def extract_syariah_data():
    with open(SYARIAH_PDF, 'rb') as f:
        data = f.read()
    result = {'PRIA': {}, 'WANITA': {}}
    for age in range(86):
        pria_stream = 15 + age * 2
        wanita_stream = 16 + age * 2
        result['PRIA'][age] = extract_numbers_from_stream(data, pria_stream)[:13]
        result['WANITA'][age] = extract_numbers_from_stream(data, wanita_stream)[:13]
    return result


def extract_konvensional_data():
    with open(KONV_PDF, 'rb') as f:
        data = f.read()
    result = {'PRIA': {}, 'WANITA': {}}
    for age in range(86):
        stream_idx = 8 if age == 0 else 10 + age
        all_nums = extract_numbers_from_stream(data, stream_idx)
        if len(all_nums) >= 60:
            pria_rs = all_nums[0:15]
            wanita_rs = all_nums[30:45]
        else:
            # Ages 80-85: only PRIA RS(15) + WANITA RS(15), no RJ
            pria_rs = all_nums[0:15]
            wanita_rs = all_nums[15:30]
        result['PRIA'][age] = [pria_rs[i] for i in KONV_INDICES]
        result['WANITA'][age] = [wanita_rs[i] for i in KONV_INDICES]
    return result


def xml_escape(s):
    return s.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;").replace('"', "&quot;")


def col_letter(col_idx):
    result = ""
    idx = col_idx
    while True:
        result = chr(65 + idx % 26) + result
        idx = idx // 26 - 1
        if idx < 0:
            break
    return result


def build_shared_strings(strings):
    ss = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    ss += f'<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"'
    ss += f' count="{len(strings)}" uniqueCount="{len(strings)}">'
    for s in strings:
        escaped = xml_escape(s)
        ss += f'<si><t>{escaped}</t></si>'
    ss += '</sst>'
    return ss


def build_styles_xml():
    """Build styles.xml matching reference file exactly."""
    s = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    s += '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"'
    s += ' xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac"'
    s += ' xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006">'
    s += '<numFmts count="1">'
    s += '<numFmt numFmtId="164" formatCode="&quot;Rp&quot;#,##0"/>'
    s += '</numFmts>'
    # fonts
    s += '<fonts count="25">'
    s += '<font><sz val="11.0"/><color rgb="FF000000"/><name val="Arial"/><scheme val="minor"/></font>'
    s += '<font><color theme="1"/><name val="Arial"/><scheme val="minor"/></font>'
    s += '<font><sz val="10.0"/><color rgb="FF000000"/><name val="Arial"/></font>'
    s += '<font><b/><sz val="11.0"/><color rgb="FFFFFFFF"/><name val="Arial"/></font>'
    s += '<font/>'
    s += '<font><b/><sz val="13.0"/><color rgb="FFFFFFFF"/><name val="Arial"/></font>'
    s += '<font><sz val="11.0"/><color theme="1"/><name val="Calibri"/></font>'
    s += '<font><sz val="10.0"/><color rgb="FFFFFFFF"/><name val="Arial"/></font>'
    s += '<font><i/><sz val="9.0"/><color rgb="FF7F7F7F"/><name val="Arial"/></font>'
    s += '<font><b/><sz val="10.0"/><color rgb="FF1E7145"/><name val="Arial"/></font>'
    s += '<font><b/><sz val="11.0"/><color rgb="FF1E7145"/><name val="Arial"/></font>'
    s += '<font><b/><sz val="14.0"/><color rgb="FF1E7145"/><name val="Arial"/></font>'
    s += '<font><i/><sz val="8.0"/><color rgb="FF7F7F7F"/><name val="Arial"/></font>'
    s += '<font><b/><sz val="13.0"/><color rgb="FF000000"/><name val="Arial"/></font>'
    s += '<font><b/><sz val="9.0"/><color rgb="FF1E7145"/><name val="Arial"/></font>'
    s += '<font><b/><sz val="10.0"/><color rgb="FFFFFFFF"/><name val="Arial"/></font>'
    s += '<font><b/><sz val="9.0"/><color rgb="FFFFFFFF"/><name val="Arial"/></font>'
    s += '<font><sz val="8.0"/><color rgb="FF333333"/><name val="Arial"/></font>'
    s += '<font><sz val="9.0"/><color rgb="FF555555"/><name val="Arial"/></font>'
    s += '<font><b/><i/><sz val="9.0"/><color rgb="FF1E7145"/><name val="Arial"/></font>'
    s += '<font><i/><sz val="8.0"/><color rgb="FF5D4037"/><name val="Arial"/></font>'
    s += '<font><b/><sz val="14.0"/><color rgb="FFFFFFFF"/><name val="Arial"/></font>'
    s += '<font><b/><i/><sz val="9.0"/><color rgb="FFC0392B"/><name val="Arial"/></font>'
    s += '<font><b/><sz val="10.0"/><color rgb="FFC0392B"/><name val="Arial"/></font>'
    s += '<font><i/><sz val="7.0"/><color rgb="FF7F7F7F"/><name val="Arial"/></font>'
    s += '</fonts>'
    # fills
    s += '<fills count="19">'
    s += '<fill><patternFill patternType="none"/></fill>'
    s += '<fill><patternFill patternType="lightGray"/></fill>'
    fills_data = [
        "FFF5F5F5", "FF1E7145", "FFE8F5EE", "FFFFFFFF", "FFFFD000",
        "FF27AE60", "FF1E8449", "FFF39C12", "FFE74C3C", "FFE8F8EF",
        "FFD5F5E3", "FFFDFEFE", "FFFFFDE7", "FFFF9800", "FF2ECC71",
        "FFE67E22", "FFFFF3CD"
    ]
    for fc in fills_data:
        s += f'<fill><patternFill patternType="solid"><fgColor rgb="{fc}"/>'
        s += f'<bgColor rgb="{fc}"/></patternFill></fill>'
    s += '</fills>'
    # borders
    s += '<borders count="10">'
    s += '<border/>'
    s += '<border><left/><right/><top/><bottom/></border>'
    s += '<border><left/><top/></border>'
    s += '<border><top/></border>'
    s += '<border><left/><top/><bottom/></border>'
    s += '<border><top/><bottom/></border>'
    s += '<border><left/></border>'
    s += '<border><left style="thin"><color rgb="FFCCCCCC"/></left>'
    s += '<right style="thin"><color rgb="FFCCCCCC"/></right>'
    s += '<top style="thin"><color rgb="FFCCCCCC"/></top>'
    s += '<bottom style="thin"><color rgb="FFCCCCCC"/></bottom></border>'
    s += '<border><left style="thin"><color rgb="FFCCCCCC"/></left>'
    s += '<top style="thin"><color rgb="FFCCCCCC"/></top>'
    s += '<bottom style="thin"><color rgb="FFCCCCCC"/></bottom></border>'
    s += '<border><top style="thin"><color rgb="FFCCCCCC"/></top>'
    s += '<bottom style="thin"><color rgb="FFCCCCCC"/></bottom></border>'
    s += '</borders>'
    # cellStyleXfs
    s += '<cellStyleXfs count="1">'
    s += '<xf borderId="0" fillId="0" fontId="0" numFmtId="0"'
    s += ' applyAlignment="1" applyFont="1"/>'
    s += '</cellStyleXfs>'
    # cellXfs - 47 styles exactly matching reference
    xfs = [
        # 0
        '<xf borderId="0" fillId="0" fontId="0" numFmtId="0" xfId="0" applyAlignment="1" applyFont="1"><alignment readingOrder="0" shrinkToFit="0" vertical="bottom" wrapText="0"/></xf>',
        # 1
        '<xf borderId="0" fillId="0" fontId="1" numFmtId="0" xfId="0" applyFont="1"/>',
        # 2
        '<xf borderId="1" fillId="2" fontId="2" numFmtId="0" xfId="0" applyAlignment="1" applyBorder="1" applyFill="1" applyFont="1"><alignment horizontal="center" shrinkToFit="0" vertical="center" wrapText="1"/></xf>',
        # 3
        '<xf borderId="1" fillId="3" fontId="2" numFmtId="0" xfId="0" applyAlignment="1" applyBorder="1" applyFill="1" applyFont="1"><alignment horizontal="center" shrinkToFit="0" vertical="center" wrapText="1"/></xf>',
        # 4
        '<xf borderId="2" fillId="2" fontId="3" numFmtId="0" xfId="0" applyAlignment="1" applyBorder="1" applyFont="1"><alignment horizontal="center" shrinkToFit="0" vertical="center" wrapText="1"/></xf>',
        # 5
        '<xf borderId="3" fillId="0" fontId="4" numFmtId="0" xfId="0" applyBorder="1" applyFont="1"/>',
        # 6
        '<xf borderId="4" fillId="3" fontId="5" numFmtId="0" xfId="0" applyAlignment="1" applyBorder="1" applyFont="1"><alignment horizontal="center" shrinkToFit="0" vertical="center" wrapText="1"/></xf>',
        # 7
        '<xf borderId="5" fillId="0" fontId="4" numFmtId="0" xfId="0" applyBorder="1" applyFont="1"/>',
        # 8
        '<xf borderId="1" fillId="2" fontId="6" numFmtId="0" xfId="0" applyAlignment="1" applyBorder="1" applyFont="1"><alignment shrinkToFit="0" vertical="bottom" wrapText="0"/></xf>',
        # 9
        '<xf borderId="6" fillId="0" fontId="4" numFmtId="0" xfId="0" applyBorder="1" applyFont="1"/>',
        # 10
        '<xf borderId="4" fillId="3" fontId="7" numFmtId="0" xfId="0" applyAlignment="1" applyBorder="1" applyFont="1"><alignment horizontal="center" shrinkToFit="0" vertical="center" wrapText="1"/></xf>',
        # 11
        '<xf borderId="4" fillId="2" fontId="8" numFmtId="0" xfId="0" applyAlignment="1" applyBorder="1" applyFont="1"><alignment horizontal="right" shrinkToFit="0" vertical="center" wrapText="1"/></xf>',
        # 12
        '<xf borderId="1" fillId="4" fontId="9" numFmtId="0" xfId="0" applyAlignment="1" applyBorder="1" applyFill="1" applyFont="1"><alignment horizontal="center" shrinkToFit="0" vertical="center" wrapText="1"/></xf>',
        # 13
        '<xf borderId="7" fillId="5" fontId="10" numFmtId="0" xfId="0" applyAlignment="1" applyBorder="1" applyFill="1" applyFont="1"><alignment horizontal="center" shrinkToFit="0" vertical="center" wrapText="1"/></xf>',
        # 14
        '<xf borderId="7" fillId="5" fontId="11" numFmtId="0" xfId="0" applyAlignment="1" applyBorder="1" applyFont="1"><alignment horizontal="center" shrinkToFit="0" vertical="center" wrapText="1"/></xf>',
        # 15
        '<xf borderId="0" fillId="0" fontId="12" numFmtId="0" xfId="0" applyAlignment="1" applyFont="1"><alignment horizontal="left" shrinkToFit="0" vertical="center" wrapText="1"/></xf>',
        # 16
        '<xf borderId="4" fillId="6" fontId="13" numFmtId="0" xfId="0" applyAlignment="1" applyBorder="1" applyFill="1" applyFont="1"><alignment horizontal="center" shrinkToFit="0" vertical="center" wrapText="1"/></xf>',
        # 17
        '<xf borderId="1" fillId="4" fontId="14" numFmtId="0" xfId="0" applyAlignment="1" applyBorder="1" applyFont="1"><alignment horizontal="right" shrinkToFit="0" vertical="center" wrapText="1"/></xf>',
        # 18
        '<xf borderId="7" fillId="7" fontId="15" numFmtId="164" xfId="0" applyAlignment="1" applyBorder="1" applyFill="1" applyFont="1" applyNumberFormat="1"><alignment horizontal="center" shrinkToFit="0" vertical="center" wrapText="1"/></xf>',
        # 19
        '<xf borderId="7" fillId="8" fontId="15" numFmtId="164" xfId="0" applyAlignment="1" applyBorder="1" applyFill="1" applyFont="1" applyNumberFormat="1"><alignment horizontal="center" shrinkToFit="0" vertical="center" wrapText="1"/></xf>',
        # 20
        '<xf borderId="7" fillId="9" fontId="15" numFmtId="164" xfId="0" applyAlignment="1" applyBorder="1" applyFill="1" applyFont="1" applyNumberFormat="1"><alignment horizontal="center" shrinkToFit="0" vertical="center" wrapText="1"/></xf>',
        # 21
        '<xf borderId="7" fillId="10" fontId="15" numFmtId="164" xfId="0" applyAlignment="1" applyBorder="1" applyFill="1" applyFont="1" applyNumberFormat="1"><alignment horizontal="center" shrinkToFit="0" vertical="center" wrapText="1"/></xf>',
        # 22
        '<xf borderId="7" fillId="7" fontId="16" numFmtId="164" xfId="0" applyAlignment="1" applyBorder="1" applyFont="1" applyNumberFormat="1"><alignment horizontal="center" shrinkToFit="0" vertical="center" wrapText="1"/></xf>',
        # 23
        '<xf borderId="7" fillId="8" fontId="16" numFmtId="164" xfId="0" applyAlignment="1" applyBorder="1" applyFont="1" applyNumberFormat="1"><alignment horizontal="center" shrinkToFit="0" vertical="center" wrapText="1"/></xf>',
        # 24
        '<xf borderId="7" fillId="9" fontId="16" numFmtId="164" xfId="0" applyAlignment="1" applyBorder="1" applyFont="1" applyNumberFormat="1"><alignment horizontal="center" shrinkToFit="0" vertical="center" wrapText="1"/></xf>',
        # 25
        '<xf borderId="7" fillId="10" fontId="16" numFmtId="164" xfId="0" applyAlignment="1" applyBorder="1" applyFont="1" applyNumberFormat="1"><alignment horizontal="center" shrinkToFit="0" vertical="center" wrapText="1"/></xf>',
        # 26
        '<xf borderId="7" fillId="7" fontId="15" numFmtId="0" xfId="0" applyAlignment="1" applyBorder="1" applyFont="1"><alignment horizontal="center" shrinkToFit="0" vertical="center" wrapText="1"/></xf>',
        # 27
        '<xf borderId="7" fillId="8" fontId="15" numFmtId="0" xfId="0" applyAlignment="1" applyBorder="1" applyFont="1"><alignment horizontal="center" shrinkToFit="0" vertical="center" wrapText="1"/></xf>',
        # 28
        '<xf borderId="7" fillId="9" fontId="15" numFmtId="0" xfId="0" applyAlignment="1" applyBorder="1" applyFont="1"><alignment horizontal="center" shrinkToFit="0" vertical="center" wrapText="1"/></xf>',
        # 29
        '<xf borderId="7" fillId="10" fontId="15" numFmtId="0" xfId="0" applyAlignment="1" applyBorder="1" applyFont="1"><alignment horizontal="center" shrinkToFit="0" vertical="center" wrapText="1"/></xf>',
        # 30
        '<xf borderId="7" fillId="11" fontId="17" numFmtId="0" xfId="0" applyAlignment="1" applyBorder="1" applyFill="1" applyFont="1"><alignment horizontal="center" shrinkToFit="0" vertical="center" wrapText="1"/></xf>',
        # 31
        '<xf borderId="7" fillId="12" fontId="17" numFmtId="0" xfId="0" applyAlignment="1" applyBorder="1" applyFill="1" applyFont="1"><alignment horizontal="center" shrinkToFit="0" vertical="center" wrapText="1"/></xf>',
        # 32
        '<xf borderId="8" fillId="13" fontId="18" numFmtId="0" xfId="0" applyAlignment="1" applyBorder="1" applyFill="1" applyFont="1"><alignment horizontal="center" shrinkToFit="0" vertical="center" wrapText="1"/></xf>',
        # 33
        '<xf borderId="9" fillId="0" fontId="4" numFmtId="0" xfId="0" applyBorder="1" applyFont="1"/>',
        # 34
        '<xf borderId="1" fillId="4" fontId="19" numFmtId="0" xfId="0" applyAlignment="1" applyBorder="1" applyFont="1"><alignment horizontal="right" shrinkToFit="0" vertical="center" wrapText="1"/></xf>',
        # 35
        '<xf borderId="8" fillId="14" fontId="20" numFmtId="0" xfId="0" applyAlignment="1" applyBorder="1" applyFill="1" applyFont="1"><alignment horizontal="center" shrinkToFit="0" vertical="center" wrapText="1"/></xf>',
        # 36
        '<xf borderId="4" fillId="15" fontId="21" numFmtId="0" xfId="0" applyAlignment="1" applyBorder="1" applyFill="1" applyFont="1"><alignment horizontal="center" shrinkToFit="0" vertical="center" wrapText="1"/></xf>',
        # 37
        '<xf borderId="8" fillId="16" fontId="3" numFmtId="164" xfId="0" applyAlignment="1" applyBorder="1" applyFill="1" applyFont="1" applyNumberFormat="1"><alignment horizontal="center" shrinkToFit="0" vertical="center" wrapText="1"/></xf>',
        # 38
        '<xf borderId="8" fillId="17" fontId="3" numFmtId="164" xfId="0" applyAlignment="1" applyBorder="1" applyFill="1" applyFont="1" applyNumberFormat="1"><alignment horizontal="center" shrinkToFit="0" vertical="center" wrapText="1"/></xf>',
        # 39
        '<xf borderId="8" fillId="16" fontId="16" numFmtId="164" xfId="0" applyAlignment="1" applyBorder="1" applyFont="1" applyNumberFormat="1"><alignment horizontal="center" shrinkToFit="0" vertical="center" wrapText="1"/></xf>',
        # 40
        '<xf borderId="8" fillId="17" fontId="16" numFmtId="164" xfId="0" applyAlignment="1" applyBorder="1" applyFont="1" applyNumberFormat="1"><alignment horizontal="center" shrinkToFit="0" vertical="center" wrapText="1"/></xf>',
        # 41
        '<xf borderId="8" fillId="16" fontId="15" numFmtId="0" xfId="0" applyAlignment="1" applyBorder="1" applyFont="1"><alignment horizontal="center" shrinkToFit="0" vertical="center" wrapText="1"/></xf>',
        # 42
        '<xf borderId="8" fillId="17" fontId="15" numFmtId="0" xfId="0" applyAlignment="1" applyBorder="1" applyFont="1"><alignment horizontal="center" shrinkToFit="0" vertical="center" wrapText="1"/></xf>',
        # 43
        '<xf borderId="8" fillId="11" fontId="17" numFmtId="0" xfId="0" applyAlignment="1" applyBorder="1" applyFont="1"><alignment horizontal="center" shrinkToFit="0" vertical="center" wrapText="1"/></xf>',
        # 44
        '<xf borderId="1" fillId="18" fontId="22" numFmtId="0" xfId="0" applyAlignment="1" applyBorder="1" applyFill="1" applyFont="1"><alignment horizontal="right" shrinkToFit="0" vertical="center" wrapText="1"/></xf>',
        # 45
        '<xf borderId="8" fillId="18" fontId="23" numFmtId="0" xfId="0" applyAlignment="1" applyBorder="1" applyFont="1"><alignment horizontal="center" shrinkToFit="0" vertical="center" wrapText="1"/></xf>',
        # 46
        '<xf borderId="4" fillId="2" fontId="24" numFmtId="0" xfId="0" applyAlignment="1" applyBorder="1" applyFont="1"><alignment horizontal="left" shrinkToFit="0" vertical="center" wrapText="1"/></xf>',
    ]
    s += f'<cellXfs count="{len(xfs)}">'
    for xf in xfs:
        s += xf
    s += '</cellXfs>'
    s += '<cellStyles count="1"><cellStyle xfId="0" name="Normal" builtinId="0"/></cellStyles>'
    s += '<dxfs count="0"/>'
    s += '</styleSheet>'
    return s


def build_database_sheet(data, ss_map):
    """Build XML for a database sheet."""
    s = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    s += '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"'
    s += ' xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
    s += '<sheetPr><pageSetUpPr/></sheetPr>'
    s += '<sheetViews><sheetView workbookViewId="0"/></sheetViews>'
    s += '<sheetFormatPr customHeight="1" defaultColWidth="12.63" defaultRowHeight="15.0"/>'
    s += '<cols><col customWidth="1" min="1" max="15" width="7.63"/></cols>'
    s += '<sheetData>'

    # Row 1: headers
    headers = ["Gender", "Usia", "Diamond", "Ruby", "Emerald", "Topaz",
               "Topaz ID", "Jade", "Jade ID", "Sapphire", "Diamond Smart",
               "Ruby Smart", "Emerald Smart", "Topaz Smart", "Jade Smart"]
    s += '<row r="1">'
    for ci, h in enumerate(headers):
        ref = col_letter(ci) + "1"
        si = ss_map[h]
        s += f'<c r="{ref}" s="1" t="s"><v>{si}</v></c>'
    s += '</row>'

    # Rows 2-87: PRIA ages 0-85
    pria_si = ss_map["PRIA"]
    wanita_si = ss_map["WANITA"]
    row_num = 2
    for age in range(86):
        ht = ' ht="15.75" customHeight="1"' if row_num >= 21 else ''
        s += f'<row r="{row_num}"{ht}>'
        s += f'<c r="A{row_num}" s="1" t="s"><v>{pria_si}</v></c>'
        s += f'<c r="B{row_num}" s="1"><v>{float(age)}</v></c>'
        for ci, val in enumerate(data['PRIA'][age]):
            ref = col_letter(ci + 2) + str(row_num)
            s += f'<c r="{ref}" s="1"><v>{val}</v></c>'
        s += '</row>'
        row_num += 1

    # Rows 88-173: WANITA ages 0-85
    for age in range(86):
        s += f'<row r="{row_num}" ht="15.75" customHeight="1">'
        s += f'<c r="A{row_num}" s="1" t="s"><v>{wanita_si}</v></c>'
        s += f'<c r="B{row_num}" s="1"><v>{float(age)}</v></c>'
        for ci, val in enumerate(data['WANITA'][age]):
            ref = col_letter(ci + 2) + str(row_num)
            s += f'<c r="{ref}" s="1"><v>{val}</v></c>'
        s += '</row>'
        row_num += 1

    # Empty rows 174-1000
    for r in range(174, 1001):
        s += f'<row r="{r}" ht="15.75" customHeight="1"/>'

    s += '</sheetData>'
    s += '<printOptions/>'
    s += '<pageMargins bottom="1.0" footer="0.0" header="0.0" left="0.75" right="0.75" top="1.0"/>'
    s += '<pageSetup paperSize="9" orientation="portrait"/>'
    s += '</worksheet>'
    return s


def build_calculator_sheet(db_sheet_name, ss_map):
    """Build XML for a calculator sheet referencing db_sheet_name."""
    # Determine formula prefix
    if ' ' in db_sheet_name:
        db_ref = f"'{db_sheet_name}'!"
    else:
        db_ref = f"{db_sheet_name}!"

    def ss(text):
        return ss_map[text]

    def formula_col(col):
        """Build the array formula for a given column letter."""
        f = f'IFERROR(ROUND(INDEX({db_ref}${col}:${col},'
        f += f'MATCH(1,({db_ref}$A:$A=IF($C$6="Pria","PRIA","WANITA"))'
        f += f'*({db_ref}$B:$B=$C$7),0))*1.12,-2),0)'
        return f

    s = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    s += '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"'
    s += ' xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
    s += '<sheetPr><pageSetUpPr fitToPage="1"/></sheetPr>'
    s += '<sheetViews><sheetView workbookViewId="0">'
    s += '<pane xSplit="2.0" ySplit="7.0" topLeftCell="C8"'
    s += ' activePane="bottomRight" state="frozen"/>'
    s += '<selection activeCell="C1" sqref="C1" pane="topRight"/>'
    s += '<selection activeCell="A8" sqref="A8" pane="bottomLeft"/>'
    s += '<selection activeCell="C8" sqref="C8" pane="bottomRight"/>'
    s += '</sheetView></sheetViews>'
    s += '<sheetFormatPr customHeight="1" defaultColWidth="12.63" defaultRowHeight="15.0"/>'
    s += '<cols>'
    s += '<col customWidth="1" min="1" max="1" width="3.5"/>'
    s += '<col customWidth="1" min="2" max="2" width="19.25"/>'
    s += '<col customWidth="1" min="3" max="3" width="15.75"/>'
    s += '<col customWidth="1" min="4" max="4" width="19.25"/>'
    s += '<col customWidth="1" min="5" max="5" width="15.75"/>'
    s += '<col customWidth="1" min="6" max="6" width="19.25"/>'
    s += '<col customWidth="1" min="7" max="7" width="3.5"/>'
    s += '</cols>'
    s += '<sheetData>'

    # Row 1: spacer
    s += '<row r="1" ht="6.0" customHeight="1">'
    s += '<c r="A1" s="2"/><c r="B1" s="3"/><c r="C1" s="3"/>'
    s += '<c r="D1" s="3"/><c r="E1" s="3"/><c r="F1" s="3"/><c r="G1" s="2"/>'
    s += '</row>'

    # Row 2: emoji + title
    emoji_str = "\U0001f3e6 Manulife"
    s += '<row r="2" ht="27.75" customHeight="1">'
    s += f'<c r="A2" s="4" t="s"><v>{ss(emoji_str)}</v></c>'
    s += '<c r="B2" s="5"/>'
    s += f'<c r="C2" s="6" t="s"><v>{ss("Mi Ultimate Healthcare (MiUHC)")}</v></c>'
    s += '<c r="D2" s="7"/><c r="E2" s="7"/><c r="F2" s="7"/><c r="G2" s="8"/>'
    s += '</row>'

    # Row 3: subtitle
    s += '<row r="3" ht="21.75" customHeight="1">'
    s += '<c r="A3" s="9"/>'
    s += f'<c r="C3" s="10" t="s"><v>{ss("PT Asuransi Jiwa Manulife Indonesia")}</v></c>'
    s += '<c r="D3" s="7"/><c r="E3" s="7"/><c r="F3" s="7"/><c r="G3" s="8"/>'
    s += '</row>'

    # Row 4: agent info with hyperlink
    s += '<row r="4" ht="18.0" customHeight="1">'
    s += f'<c r="A4" s="11" t="s"><v>{ss("PETRUS JAKUB MANULIFE 087781896087")}</v></c>'
    s += '<c r="B4" s="7"/><c r="C4" s="7"/><c r="D4" s="7"/>'
    s += '<c r="E4" s="7"/><c r="F4" s="7"/><c r="G4" s="8"/>'
    s += '</row>'

    # Row 5: spacer
    s += '<row r="5" ht="7.5" customHeight="1">'
    s += '<c r="A5" s="8"/><c r="G5" s="8"/>'
    s += '</row>'

    # Row 6: Jenis Kelamin
    s += '<row r="6" ht="24.0" customHeight="1">'
    s += '<c r="A6" s="8"/>'
    s += f'<c r="B6" s="12" t="s"><v>{ss("Jenis Kelamin")}</v></c>'
    s += f'<c r="C6" s="13" t="s"><v>{ss("Pria")}</v></c>'
    s += '<c r="G6" s="8"/>'
    s += '</row>'

    # Row 7: Usia
    s += '<row r="7" ht="24.0" customHeight="1">'
    s += '<c r="A7" s="8"/>'
    s += f'<c r="B7" s="12" t="s"><v>{ss("Usia")}</v></c>'
    s += '<c r="C7" s="14"><v>31.0</v></c>'
    s += f'<c r="D7" s="15" t="s"><v>{ss("*Usia setelah ulang tahun terakhir")}</v></c>'
    s += '<c r="G7" s="8"/>'
    s += '</row>'

    # Row 8: spacer
    s += '<row r="8" ht="9.75" customHeight="1">'
    s += '<c r="A8" s="8"/><c r="G8" s="8"/>'
    s += '</row>'

    # Row 9: Most Wanted header
    s += '<row r="9" ht="25.5" customHeight="1">'
    s += '<c r="A9" s="8"/>'
    s += f'<c r="B9" s="16" t="s"><v>{ss("Most Wanted  [ No Deductible ]")}</v></c>'
    s += '<c r="C9" s="7"/><c r="D9" s="7"/><c r="E9" s="7"/><c r="F9" s="7"/>'
    s += '<c r="G9" s="8"/>'
    s += '</row>'

    # Row 10: Premi Tahunan with array formulas
    fc10 = xml_escape(formula_col("I"))
    fd10 = xml_escape(formula_col("H"))
    fe10 = xml_escape(formula_col("G"))
    ff10 = xml_escape(formula_col("F"))
    s += '<row r="10" ht="21.75" customHeight="1">'
    s += '<c r="A10" s="8"/>'
    s += f'<c r="B10" s="17" t="s"><v>{ss("Premi Tahunan")}</v></c>'
    s += f'<c r="C10" s="18"><f t="array" ref="C10">{fc10}</f><v>0</v></c>'
    s += f'<c r="D10" s="19"><f t="array" ref="D10">{fd10}</f><v>0</v></c>'
    s += f'<c r="E10" s="20"><f t="array" ref="E10">{fe10}</f><v>0</v></c>'
    s += f'<c r="F10" s="21"><f t="array" ref="F10">{ff10}</f><v>0</v></c>'
    s += '<c r="G10" s="8"/>'
    s += '</row>'

    # Row 11: Premi Bulanan shared formula
    monthly_f = 'IFERROR(ROUND(C10*1.14/12,-2),0)'
    s += '<row r="11" ht="19.5" customHeight="1">'
    s += '<c r="A11" s="8"/>'
    s += f'<c r="B11" s="17" t="s"><v>{ss("Premi Bulanan")}</v></c>'
    s += f'<c r="C11" s="22"><f t="shared" ref="C11:F11" si="1">{monthly_f}</f><v>0</v></c>'
    s += '<c r="D11" s="23"><f t="shared" si="1"/><v>0</v></c>'
    s += '<c r="E11" s="24"><f t="shared" si="1"/><v>0</v></c>'
    s += '<c r="F11" s="25"><f t="shared" si="1"/><v>0</v></c>'
    s += '<c r="G11" s="8"/>'
    s += '</row>'

    # Row 12: Plan labels
    s += '<row r="12" ht="19.5" customHeight="1">'
    s += '<c r="A12" s="8"/>'
    s += f'<c r="B12" s="17" t="s"><v>{ss("Plan")}</v></c>'
    s += f'<c r="C12" s="26" t="s"><v>{ss("Jade ID")}</v></c>'
    s += f'<c r="D12" s="27" t="s"><v>{ss("Jade")}</v></c>'
    s += f'<c r="E12" s="28" t="s"><v>{ss("Topaz ID")}</v></c>'
    s += f'<c r="F12" s="29" t="s"><v>{ss("Topaz")}</v></c>'
    s += '<c r="G12" s="8"/>'
    s += '</row>'

    # Row 13: Wilayah
    s += '<row r="13" ht="30.0" customHeight="1">'
    s += '<c r="A13" s="8"/>'
    s += f'<c r="B13" s="17" t="s"><v>{ss("Wilayah")}</v></c>'
    s += f'<c r="C13" s="30" t="s"><v>{ss("Indonesia")}</v></c>'
    s += f'<c r="D13" s="31" t="s"><v>{ss("Seluruh Asia (kec." + chr(10) + "JPN/HKG/SGP)")}</v></c>'
    s += f'<c r="E13" s="30" t="s"><v>{ss("Indonesia")}</v></c>'
    s += f'<c r="F13" s="31" t="s"><v>{ss("Seluruh Asia (kec." + chr(10) + "JPN/HKG/SGP)")}</v></c>'
    s += '<c r="G13" s="8"/>'
    s += '</row>'

    # Row 14: Kamar
    s += '<row r="14" ht="18.0" customHeight="1">'
    s += '<c r="A14" s="8"/>'
    s += f'<c r="B14" s="17" t="s"><v>{ss("Kamar")}</v></c>'
    s += f'<c r="C14" s="32" t="s"><v>{ss("2 Bed")}</v></c>'
    s += '<c r="D14" s="33"/>'
    s += f'<c r="E14" s="32" t="s"><v>{ss("1 Bed")}</v></c>'
    s += '<c r="F14" s="33"/><c r="G14" s="8"/>'
    s += '</row>'

    # Row 15: Limit Kartu
    s += '<row r="15" ht="18.0" customHeight="1">'
    s += '<c r="A15" s="8"/>'
    s += f'<c r="B15" s="17" t="s"><v>{ss("Limit Kartu")}</v></c>'
    s += f'<c r="C15" s="32" t="s"><v>{ss("10 M")}</v></c>'
    s += '<c r="D15" s="33"/>'
    s += f'<c r="E15" s="32" t="s"><v>{ss("18 M")}</v></c>'
    s += '<c r="F15" s="33"/><c r="G15" s="8"/>'
    s += '</row>'

    # Row 16: Asuransi Jiwa
    s += '<row r="16" ht="18.0" customHeight="1">'
    s += '<c r="A16" s="8"/>'
    s += f'<c r="B16" s="17" t="s"><v>{ss("Asuransi Jiwa")}</v></c>'
    s += f'<c r="C16" s="32" t="s"><v>{ss("30 Juta")}</v></c>'
    s += '<c r="D16" s="33"/><c r="E16" s="33"/><c r="F16" s="33"/><c r="G16" s="8"/>'
    s += '</row>'

    # Row 17: Rawat Inap
    s += '<row r="17" ht="18.0" customHeight="1">'
    s += '<c r="A17" s="8"/>'
    s += f'<c r="B17" s="17" t="s"><v>{ss("Rawat Inap")}</v></c>'
    s += f'<c r="C17" s="32" t="s"><v>{ss("Sesuai Tagihan")}</v></c>'
    s += '<c r="D17" s="33"/><c r="E17" s="33"/><c r="F17" s="33"/><c r="G17" s="8"/>'
    s += '</row>'

    # Row 18: No-Claim BONUS
    bonus_text = "Setahun tanpa klaim? LIMIT tahunan kartu Anda naik 10% per tahun hingga 50%"
    s += '<row r="18" ht="27.75" customHeight="1">'
    s += '<c r="A18" s="8"/>'
    s += f'<c r="B18" s="34" t="s"><v>{ss("No-Claim BONUS")}</v></c>'
    s += f'<c r="C18" s="35" t="s"><v>{ss(bonus_text)}</v></c>'
    s += '<c r="D18" s="33"/><c r="E18" s="33"/><c r="F18" s="33"/><c r="G18" s="8"/>'
    s += '</row>'

    # Row 19: No-Claim DISCOUNT
    disc_text = "Setahun tanpa klaim? Dapatkan DISCOUNT 10% di tahun berikutnya, hingga 15%"
    s += '<row r="19" ht="27.75" customHeight="1">'
    s += '<c r="A19" s="8"/>'
    s += f'<c r="B19" s="34" t="s"><v>{ss("No-Claim DISCOUNT")}</v></c>'
    s += f'<c r="C19" s="35" t="s"><v>{ss(disc_text)}</v></c>'
    s += '<c r="D19" s="33"/><c r="E19" s="33"/><c r="F19" s="33"/><c r="G19" s="8"/>'
    s += '</row>'

    # Row 20: spacer
    s += '<row r="20" ht="12.0" customHeight="1">'
    s += '<c r="A20" s="8"/><c r="G20" s="8"/>'
    s += '</row>'

    # Row 21: HEMAT 40% header
    s += '<row r="21" ht="25.5" customHeight="1">'
    s += '<c r="A21" s="8"/>'
    s += f'<c r="B21" s="36" t="s"><v>{ss("HEMAT 40%")}</v></c>'
    s += '<c r="C21" s="7"/><c r="D21" s="7"/><c r="E21" s="7"/><c r="F21" s="7"/>'
    s += '<c r="G21" s="8"/>'
    s += '</row>'

    # Row 22: Smart Premi Tahunan
    fo22 = xml_escape(formula_col("O"))
    fn22 = xml_escape(formula_col("N"))
    s += '<row r="22" ht="21.75" customHeight="1">'
    s += '<c r="A22" s="8"/>'
    s += f'<c r="B22" s="17" t="s"><v>{ss("Premi Tahunan")}</v></c>'
    s += f'<c r="C22" s="37"><f t="array" ref="C22">{fo22}</f><v>0</v></c>'
    s += '<c r="D22" s="33"/>'
    s += f'<c r="E22" s="38"><f t="array" ref="E22">{fn22}</f><v>0</v></c>'
    s += '<c r="F22" s="33"/><c r="G22" s="8"/>'
    s += '</row>'

    # Row 23: Smart Premi Bulanan
    s += '<row r="23" ht="19.5" customHeight="1">'
    s += '<c r="A23" s="8"/>'
    s += f'<c r="B23" s="17" t="s"><v>{ss("Premi Bulanan")}</v></c>'
    s += '<c r="C23" s="39"><f>IFERROR(ROUND(C22*1.14/12,-2),0)</f><v>0</v></c>'
    s += '<c r="D23" s="33"/>'
    s += '<c r="E23" s="40"><f>IFERROR(ROUND(E22*1.14/12,-2),0)</f><v>0</v></c>'
    s += '<c r="F23" s="33"/><c r="G23" s="8"/>'
    s += '</row>'

    # Row 24: Smart Plan labels
    s += '<row r="24" ht="19.5" customHeight="1">'
    s += '<c r="A24" s="8"/>'
    s += f'<c r="B24" s="17" t="s"><v>{ss("Plan")}</v></c>'
    s += f'<c r="C24" s="41" t="s"><v>{ss("Jade SMART")}</v></c>'
    s += '<c r="D24" s="33"/>'
    s += f'<c r="E24" s="42" t="s"><v>{ss("Topaz SMART")}</v></c>'
    s += '<c r="F24" s="33"/><c r="G24" s="8"/>'
    s += '</row>'

    # Row 25: Smart Wilayah
    asia_text = "Seluruh Asia (kec. JPN/ HKG / SGP)"
    s += '<row r="25" ht="30.0" customHeight="1">'
    s += '<c r="A25" s="8"/>'
    s += f'<c r="B25" s="17" t="s"><v>{ss("Wilayah")}</v></c>'
    s += f'<c r="C25" s="43" t="s"><v>{ss(asia_text)}</v></c>'
    s += '<c r="D25" s="33"/>'
    s += f'<c r="E25" s="43" t="s"><v>{ss(asia_text)}</v></c>'
    s += '<c r="F25" s="33"/><c r="G25" s="8"/>'
    s += '</row>'

    # Row 26: Smart Kamar
    s += '<row r="26" ht="18.0" customHeight="1">'
    s += '<c r="A26" s="8"/>'
    s += f'<c r="B26" s="17" t="s"><v>{ss("Kamar")}</v></c>'
    s += f'<c r="C26" s="32" t="s"><v>{ss("2 Bed")}</v></c>'
    s += '<c r="D26" s="33"/>'
    s += f'<c r="E26" s="32" t="s"><v>{ss("1 Bed")}</v></c>'
    s += '<c r="F26" s="33"/><c r="G26" s="8"/>'
    s += '</row>'

    # Row 27: Deductible
    s += '<row r="27" ht="18.0" customHeight="1">'
    s += '<c r="A27" s="8"/>'
    s += f'<c r="B27" s="44" t="s"><v>{ss("Deductible*")}</v></c>'
    s += f'<c r="C27" s="45" t="s"><v>{ss("Rp5.000.000")}</v></c>'
    s += '<c r="D27" s="33"/>'
    s += f'<c r="E27" s="45" t="s"><v>{ss("Rp8.000.000")}</v></c>'
    s += '<c r="F27" s="33"/><c r="G27" s="8"/>'
    s += '</row>'

    # Row 28: bottom spacer
    s += '<row r="28" ht="6.0" customHeight="1">'
    s += '<c r="A28" s="2"/><c r="B28" s="3"/><c r="C28" s="3"/>'
    s += '<c r="D28" s="3"/><c r="E28" s="3"/><c r="F28" s="3"/><c r="G28" s="2"/>'
    s += '</row>'

    # Row 29: footer note
    footer = '*Deductible = biaya yang ditanggung nasabah sendiri sebelum manfaat asuransi berlaku  |  Premi = RS Premi \xd7 1.12 (tahunan) \u2192 bulanan \xd7 1.14 / 12'
    s += '<row r="29" ht="15.75" customHeight="1">'
    s += f'<c r="A29" s="46" t="s"><v>{ss(footer)}</v></c>'
    s += '<c r="B29" s="7"/><c r="C29" s="7"/><c r="D29" s="7"/>'
    s += '<c r="E29" s="7"/><c r="F29" s="7"/><c r="G29" s="8"/>'
    s += '</row>'

    # Rows 30-1000: empty
    for r in range(30, 1001):
        s += f'<row r="{r}" ht="15.75" customHeight="1"/>'

    s += '</sheetData>'

    # Merge cells
    s += '<mergeCells count="28">'
    merges = [
        "A2:B3", "C2:F2", "C3:F3", "A4:F4", "D7:F7", "B9:F9",
        "E14:F14", "C14:D14", "C15:D15", "E15:F15", "C16:F16",
        "C17:F17", "C18:F18", "C19:F19", "C25:D25", "E25:F25",
        "C26:D26", "E26:F26", "C27:D27", "E27:F27", "A29:F29",
        "B21:F21", "C22:D22", "E22:F22", "C23:D23", "E23:F23",
        "C24:D24", "E24:F24"
    ]
    for m in merges:
        s += f'<mergeCell ref="{m}"/>'
    s += '</mergeCells>'

    # Data validations
    s += '<dataValidations>'
    s += '<dataValidation type="list" allowBlank="1" sqref="C6">'
    s += '<formula1>&quot;Pria,Wanita&quot;</formula1>'
    s += '</dataValidation>'
    s += '<dataValidation type="list" allowBlank="1" sqref="C7">'
    ages = ','.join(str(i) for i in range(86))
    s += f'<formula1>&quot;{ages}&quot;</formula1>'
    s += '</dataValidation>'
    s += '</dataValidations>'

    s += '<printOptions/>'
    s += '<pageMargins bottom="1.0" footer="0.0" header="0.0"'
    s += ' left="0.75" right="0.75" top="1.0"/>'
    s += '<pageSetup orientation="portrait"/>'
    s += '<hyperlinks>'
    s += '<hyperlink ref="A4" r:id="rId1"/>'
    s += '</hyperlinks>'
    s += '</worksheet>'
    return s


def build_theme_xml():
    """Build a minimal theme1.xml."""
    t = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    t += '<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"'
    t += ' name="Sheets">'
    t += '<a:themeElements>'
    t += '<a:clrScheme name="Sheets">'
    t += '<a:dk1><a:srgbClr val="000000"/></a:dk1>'
    t += '<a:lt1><a:srgbClr val="FFFFFF"/></a:lt1>'
    t += '<a:dk2><a:srgbClr val="000000"/></a:dk2>'
    t += '<a:lt2><a:srgbClr val="FFFFFF"/></a:lt2>'
    t += '<a:accent1><a:srgbClr val="4F81BD"/></a:accent1>'
    t += '<a:accent2><a:srgbClr val="C0504D"/></a:accent2>'
    t += '<a:accent3><a:srgbClr val="9BBB59"/></a:accent3>'
    t += '<a:accent4><a:srgbClr val="8064A2"/></a:accent4>'
    t += '<a:accent5><a:srgbClr val="4BACC6"/></a:accent5>'
    t += '<a:accent6><a:srgbClr val="F79646"/></a:accent6>'
    t += '<a:hlink><a:srgbClr val="0000FF"/></a:hlink>'
    t += '<a:folHlink><a:srgbClr val="0000FF"/></a:folHlink>'
    t += '</a:clrScheme>'
    t += '<a:fontScheme name="Sheets">'
    t += '<a:majorFont><a:latin typeface="Arial"/>'
    t += '<a:ea typeface="Arial"/><a:cs typeface="Arial"/></a:majorFont>'
    t += '<a:minorFont><a:latin typeface="Arial"/>'
    t += '<a:ea typeface="Arial"/><a:cs typeface="Arial"/></a:minorFont>'
    t += '</a:fontScheme>'
    t += '<a:fmtScheme name="Office">'
    t += '<a:fillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill>'
    t += '<a:solidFill><a:schemeClr val="phClr"/></a:solidFill>'
    t += '<a:solidFill><a:schemeClr val="phClr"/></a:solidFill>'
    t += '</a:fillStyleLst>'
    t += '<a:lnStyleLst><a:ln><a:solidFill><a:schemeClr val="phClr"/>'
    t += '</a:solidFill></a:ln><a:ln><a:solidFill><a:schemeClr val="phClr"/>'
    t += '</a:solidFill></a:ln><a:ln><a:solidFill><a:schemeClr val="phClr"/>'
    t += '</a:solidFill></a:ln></a:lnStyleLst>'
    t += '<a:effectStyleLst><a:effectStyle><a:effectLst/></a:effectStyle>'
    t += '<a:effectStyle><a:effectLst/></a:effectStyle>'
    t += '<a:effectStyle><a:effectLst/></a:effectStyle></a:effectStyleLst>'
    t += '<a:bgFillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill>'
    t += '<a:solidFill><a:schemeClr val="phClr"/></a:solidFill>'
    t += '<a:solidFill><a:schemeClr val="phClr"/></a:solidFill>'
    t += '</a:bgFillStyleLst>'
    t += '</a:fmtScheme>'
    t += '</a:themeElements></a:theme>'
    return t


def main():
    print("Extracting premium data from PDFs...")
    syariah_data = extract_syariah_data()
    konv_data = extract_konvensional_data()

    print(f"  Syariah PRIA age 0 Diamond: {syariah_data['PRIA'][0][0]}")
    print(f"  Konvensional PRIA age 0 Diamond: {konv_data['PRIA'][0][0]}")

    # Build shared strings table
    # Collect all unique strings needed
    all_strings = [
        "Gender", "Usia", "Diamond", "Ruby", "Emerald", "Topaz",
        "Topaz ID", "Jade", "Jade ID", "Sapphire", "Diamond Smart",
        "Ruby Smart", "Emerald Smart", "Topaz Smart", "Jade Smart",
        "PRIA", "WANITA",
        "\U0001f3e6 Manulife",
        "Mi Ultimate Healthcare (MiUHC)",
        "PT Asuransi Jiwa Manulife Indonesia",
        "PETRUS JAKUB MANULIFE 087781896087",
        "Jenis Kelamin", "Pria",
        "*Usia setelah ulang tahun terakhir",
        "Most Wanted  [ No Deductible ]",
        "Premi Tahunan", "Premi Bulanan", "Plan", "Wilayah",
        "Indonesia", "Seluruh Asia (kec.\nJPN/HKG/SGP)",
        "Kamar", "2 Bed", "1 Bed", "Limit Kartu", "10 M", "18 M",
        "Asuransi Jiwa", "30 Juta", "Rawat Inap", "Sesuai Tagihan",
        "No-Claim BONUS",
        "Setahun tanpa klaim? LIMIT tahunan kartu Anda naik 10% per tahun hingga 50%",
        "No-Claim DISCOUNT",
        "Setahun tanpa klaim? Dapatkan DISCOUNT 10% di tahun berikutnya, hingga 15%",
        "HEMAT 40%",
        "Jade SMART", "Topaz SMART",
        "Seluruh Asia (kec. JPN/ HKG / SGP)",
        "Deductible*", "Rp5.000.000", "Rp8.000.000",
        "*Deductible = biaya yang ditanggung nasabah sendiri sebelum manfaat asuransi berlaku  |  Premi = RS Premi \xd7 1.12 (tahunan) \u2192 bulanan \xd7 1.14 / 12",
    ]
    ss_map = {s: i for i, s in enumerate(all_strings)}

    print("Building XML components...")

    # Build all sheet XMLs
    db_konv_xml = build_database_sheet(konv_data, ss_map)
    calc_konv_xml = build_calculator_sheet("Database", ss_map)
    db_syariah_xml = build_database_sheet(syariah_data, ss_map)
    calc_syariah_xml = build_calculator_sheet("Database Syariah", ss_map)

    # Content Types
    ct = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    ct += '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
    ct += '<Default ContentType="application/xml" Extension="xml"/>'
    ct += '<Default ContentType="application/vnd.openxmlformats-package.relationships+xml" Extension="rels"/>'
    for i in range(1, 5):
        ct += f'<Override ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml" PartName="/xl/worksheets/sheet{i}.xml"/>'
    ct += '<Override ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml" PartName="/xl/sharedStrings.xml"/>'
    ct += '<Override ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml" PartName="/xl/styles.xml"/>'
    ct += '<Override ContentType="application/vnd.openxmlformats-officedocument.theme+xml" PartName="/xl/theme/theme1.xml"/>'
    ct += '<Override ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml" PartName="/xl/workbook.xml"/>'
    ct += '</Types>'

    # Top-level rels
    rels = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    rels += '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
    rels += '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>'
    rels += '</Relationships>'

    # Workbook
    wb = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    wb += '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"'
    wb += ' xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
    wb += '<workbookPr/><sheets>'
    wb += '<sheet state="hidden" name="Database" sheetId="1" r:id="rId4"/>'
    wb += '<sheet state="visible" name="Kalkulator MiUHC" sheetId="2" r:id="rId5"/>'
    wb += '<sheet state="hidden" name="Database Syariah" sheetId="3" r:id="rId6"/>'
    wb += '<sheet state="visible" name="Kalkulator MiUHC Syariah" sheetId="4" r:id="rId7"/>'
    wb += '</sheets><definedNames/><calcPr/></workbook>'

    # Workbook rels
    wb_rels = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    wb_rels += '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
    wb_rels += '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml"/>'
    wb_rels += '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>'
    wb_rels += '<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/>'
    wb_rels += '<Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>'
    wb_rels += '<Relationship Id="rId5" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet2.xml"/>'
    wb_rels += '<Relationship Id="rId6" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet3.xml"/>'
    wb_rels += '<Relationship Id="rId7" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet4.xml"/>'
    wb_rels += '</Relationships>'

    # Sheet rels for calculator sheets (hyperlinks)
    sheet_rels = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    sheet_rels += '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
    sheet_rels += f'<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="{WA_LINK}" TargetMode="External"/>'
    sheet_rels += '</Relationships>'

    # Shared strings
    shared_strings_xml = build_shared_strings(all_strings)
    styles_xml = build_styles_xml()
    theme_xml = build_theme_xml()

    print("Writing xlsx file...")
    with zipfile.ZipFile(OUTPUT_XLSX, 'w', zipfile.ZIP_DEFLATED) as zf:
        zf.writestr('[Content_Types].xml', ct)
        zf.writestr('_rels/.rels', rels)
        zf.writestr('xl/workbook.xml', wb)
        zf.writestr('xl/_rels/workbook.xml.rels', wb_rels)
        zf.writestr('xl/styles.xml', styles_xml)
        zf.writestr('xl/sharedStrings.xml', shared_strings_xml)
        zf.writestr('xl/theme/theme1.xml', theme_xml)
        zf.writestr('xl/worksheets/sheet1.xml', db_konv_xml)
        zf.writestr('xl/worksheets/sheet2.xml', calc_konv_xml)
        zf.writestr('xl/worksheets/sheet3.xml', db_syariah_xml)
        zf.writestr('xl/worksheets/sheet4.xml', calc_syariah_xml)
        # Hyperlink rels for calculator sheets (sheet2, sheet4)
        zf.writestr('xl/worksheets/_rels/sheet2.xml.rels', sheet_rels)
        zf.writestr('xl/worksheets/_rels/sheet4.xml.rels', sheet_rels)

    print(f"Output: {OUTPUT_XLSX}")

    # Verify
    with zipfile.ZipFile(OUTPUT_XLSX) as z:
        print(f"  Valid xlsx with {len(z.namelist())} files")
        print(f"  Contents: {z.namelist()}")


if __name__ == "__main__":
    main()
