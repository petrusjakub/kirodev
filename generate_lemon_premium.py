#!/usr/bin/env python3
"""
Generate 保费费率表.xlsx - Premium Rate Table for Lemon Overseas Insurance
Basic Sum Assured: 100,000 (基本保额10万)
Uses only Python standard library (zipfile + xml.etree.ElementTree)
"""

import zipfile
import xml.etree.ElementTree as ET
from io import BytesIO
import os

# Premium data: [age, gender, 趸交, 5年交, 10年交, 15年交, 20年交, 30年交]
# None means not available for that payment term
PREMIUM_DATA = [
    [0, "男", 92610, 19498, 10296, 7218, 5668, 4056],
    [0, "女", 88498, 18630, 9834, 6892, 5410, 3868],
    [1, "男", 93740, 19738, 10424, 7308, 5738, 4108],
    [1, "女", 89590, 18860, 9956, 6978, 5480, 3918],
    [2, "男", 94900, 19984, 10554, 7400, 5810, 4160],
    [2, "女", 90710, 19096, 10080, 7066, 5550, 3970],
    [3, "男", 96090, 20236, 10688, 7494, 5884, 4214],
    [3, "女", 91860, 19338, 10208, 7156, 5620, 4022],
    [4, "男", 97310, 20494, 10824, 7590, 5960, 4268],
    [4, "女", 93040, 19586, 10340, 7248, 5694, 4076],
    [5, "男", 98570, 20760, 10966, 7690, 6038, 4326],
    [5, "女", 94250, 19842, 10476, 7344, 5768, 4130],
    [6, "男", 99870, 21034, 11112, 7792, 6120, 4384],
    [6, "女", 95500, 20106, 10616, 7442, 5846, 4186],
    [7, "男", 101210, 21318, 11264, 7900, 6204, 4446],
    [7, "女", 96790, 20378, 10760, 7544, 5926, 4244],
    [8, "男", 102600, 21612, 11422, 8010, 6292, 4510],
    [8, "女", 98120, 20660, 10910, 7650, 6010, 4304],
    [9, "男", 104040, 21918, 11586, 8126, 6384, 4578],
    [9, "女", 99500, 20952, 11066, 7760, 6098, 4368],
    [10, "男", 105540, 22236, 11758, 8248, 6480, 4648],
    [10, "女", 100930, 21256, 11228, 7876, 6190, 4434],
    [11, "男", 107100, 22568, 11938, 8376, 6582, 4722],
    [11, "女", 102410, 21572, 11396, 7996, 6286, 4504],
    [12, "男", 108730, 22914, 12126, 8510, 6690, 4800],
    [12, "女", 103950, 21900, 11572, 8122, 6386, 4578],
    [13, "男", 110430, 23278, 12322, 8652, 6804, 4882],
    [13, "女", 105550, 22242, 11756, 8254, 6492, 4656],
    [14, "男", 112210, 23660, 12528, 8802, 6924, 4970],
    [14, "女", 107220, 22600, 11950, 8392, 6604, 4738],
    [15, "男", 114080, 24062, 12746, 8960, 7052, 5062],
    [15, "女", 108970, 22974, 12154, 8538, 6722, 4826],
    [16, "男", 116040, 24488, 12974, 9128, 7188, 5162],
    [16, "女", 110800, 23366, 12368, 8694, 6848, 4918],
    [17, "男", 118110, 24940, 13218, 9306, 7332, 5268],
    [17, "女", 112730, 23780, 12596, 8858, 6980, 5016],
    [18, "男", 120290, 25418, 13476, 9496, 7488, 5382],
    [18, "女", 114760, 24218, 12838, 9034, 7122, 5122],
    [19, "男", 122600, 25926, 13752, 9698, 7654, None],
    [19, "女", 116910, 24682, 13096, 9222, 7276, None],
    [20, "男", 125060, 26468, 14048, 9916, 7834, None],
    [20, "女", 119190, 25176, 13370, 9424, 7442, None],
    [21, "男", 127680, 27046, 14366, 10150, 8028, None],
    [21, "女", 121610, 25702, 13662, 9642, 7622, None],
    [22, "男", 130470, 27666, 14706, 10402, 8236, None],
    [22, "女", 124180, 26266, 13974, 9876, 7816, None],
    [23, "男", 133450, 28332, 15072, 10672, 8462, None],
    [23, "女", 126920, 26872, 14308, 10128, 8024, None],
    [24, "男", 136630, 29048, 15466, 10962, 8704, None],
    [24, "女", 129840, 27522, 14666, 10398, 8250, None],
    [25, "男", 140030, 29818, 15890, 11274, 8964, None],
    [25, "女", 132960, 28222, 15052, 10690, 8494, None],
    [26, "男", 143670, 30648, 16346, 11612, 9244, None],
    [26, "女", 136290, 28978, 15468, 11004, 8758, None],
    [27, "男", 147570, 31544, 16840, 11976, 9546, None],
    [27, "女", 139860, 29794, 15918, 11342, 9042, None],
    [28, "男", 151760, 32514, 17374, 12372, 9872, None],
    [28, "女", 143690, 30676, 16402, 11706, 9350, None],
    [29, "男", 156260, 33562, 17950, 12800, 10224, None],
    [29, "女", 147810, 31630, 16926, 12100, 9682, None],
    [30, "男", 161100, 34698, 18574, 13264, 10606, None],
    [30, "女", 152240, 32664, 17494, 12528, 10040, None],
    [31, "男", 166310, 35928, 19250, 13768, 11020, None],
    [31, "女", 157010, 33786, 18110, 12990, 10428, None],
    [32, "男", 171930, 37262, 19984, 14312, 11470, None],
    [32, "女", 162150, 35002, 18780, 13492, 10848, None],
    [33, "男", 178000, 38710, 20782, 14902, 11956, None],
    [33, "女", 167700, 36322, 19508, 14034, 11302, None],
    [34, "男", 184560, 40282, 21644, 15540, 12482, None],
    [34, "女", 173690, 37756, 20300, 14622, 11794, None],
    [35, "男", 191650, 41990, 22578, 16230, 13052, None],
    [35, "女", 180160, 39312, 21160, 15260, 12328, None],
    [36, "男", 199320, 43848, 23588, 16976, 13672, None],
    [36, "女", 187150, 41004, 22094, 15952, 12906, None],
    [37, "男", 207620, 45870, 24682, 17784, 14342, None],
    [37, "女", 194710, 42846, 23108, 16702, 13534, None],
    [38, "男", 216610, 48072, 25868, 18660, 15068, None],
    [38, "女", 202890, 44854, 24206, 17516, 14218, None],
    [39, "男", 226350, 50470, 27154, 19610, 15854, None],
    [39, "女", 211740, 47042, 25396, 18400, 14958, None],
    [40, "男", 236910, 53082, 28552, 20638, 16706, None],
    [40, "女", 221320, 49424, 26690, 19360, 15762, None],
    [41, "男", 248360, 55928, 30076, 21752, 17630, None],
    [41, "女", 231700, 52018, 28098, 20404, 16636, None],
    [42, "男", 260770, 59030, 31738, 22960, 18634, None],
    [42, "女", 242940, 54842, 29632, 21538, 17588, None],
    [43, "男", 274220, 62408, 33552, 24272, 19726, None],
    [43, "女", 255110, 57918, 31306, 22774, 18628, None],
    [44, "男", 288790, 66086, 35534, 25700, 20914, None],
    [44, "女", 268290, 61268, 33132, 24120, 19762, None],
    [45, "男", 304570, 70092, 37700, 27258, 22210, None],
    [45, "女", 282560, 64916, 35126, 25588, 21002, None],
    [46, "男", 321650, 74454, 40062, 28960, None, None],
    [46, "女", 298010, 68890, 37300, 27190, None, None],
    [47, "男", 340130, 79200, 42636, 30826, None, None],
    [47, "女", 314720, 73218, 39668, 28936, None, None],
    [48, "男", 360110, 84362, 45440, 32872, None, None],
    [48, "女", 332790, 77928, 42244, 30842, None, None],
    [49, "男", 381700, 89976, 48498, 35110, None, None],
    [49, "女", 352320, 83050, 45046, 32928, None, None],
    [50, "男", 405030, 96078, 51832, None, None, None],
    [50, "女", 373420, 88618, 48094, None, None, None],
    [51, "男", 430240, 102710, 55468, None, None, None],
    [51, "女", 396220, 94668, 51404, None, None, None],
    [52, "男", 457490, 109920, 59430, None, None, None],
    [52, "女", 420850, 101240, 54984, None, None, None],
    [53, "男", 487000, 117770, 63748, None, None, None],
    [53, "女", 447460, 108390, 58862, None, None, None],
    [54, "男", 518930, 126310, 68450, None, None, None],
    [54, "女", 476130, 116160, 63070, None, None, None],
    [55, "男", 553510, 135600, 73570, None, None, None],
    [55, "女", 507020, 124610, 67640, None, None, None],
]

# Column headers
HEADERS = ["年龄", "性别", "趸交", "5年交", "10年交", "15年交", "20年交", "30年交"]
TITLE = "保费费率表 - 基本保额10万"
NOTE = "注：以上保费为基本保额10万元对应的年缴保费（单位：元）"
SHEET_NAME = "保费表"


def col_letter(col_idx):
    """Convert 0-based column index to Excel column letter (A, B, C, ...)"""
    return chr(65 + col_idx)


def format_number(n):
    """Format number with thousands separator"""
    if n is None:
        return None
    s = str(n)
    result = []
    for i, ch in enumerate(reversed(s)):
        if i > 0 and i % 3 == 0:
            result.append(',')
        result.append(ch)
    return ''.join(reversed(result))


def create_xlsx(output_path):
    """Create the xlsx file using zipfile and raw XML"""

    # Shared strings table
    shared_strings = []
    ss_index_map = {}

    def get_ss_index(text):
        if text not in ss_index_map:
            ss_index_map[text] = len(shared_strings)
            shared_strings.append(text)
        return ss_index_map[text]

    # Build the sheet data
    # Row 1: Title (merged across all columns)
    # Row 2: Empty
    # Row 3: Headers
    # Row 4+: Data rows
    # Last row: Note

    rows_data = []

    # Row 1: Title
    rows_data.append([("s", TITLE)] + [("empty",)] * 7)

    # Row 2: Empty
    rows_data.append([("empty",)] * 8)

    # Row 3: Headers
    header_row = []
    for h in HEADERS:
        header_row.append(("s", h))
    rows_data.append(header_row)

    # Data rows
    for record in PREMIUM_DATA:
        age, gender = record[0], record[1]
        row = [("n", age), ("s", gender)]
        for val in record[2:]:
            if val is None:
                row.append(("s", "-"))
            else:
                row.append(("n", val))
        rows_data.append(row)

    # Empty row before note
    rows_data.append([("empty",)] * 8)

    # Note row
    rows_data.append([("s", NOTE)] + [("empty",)] * 7)

    # Pre-populate shared strings
    for row in rows_data:
        for cell in row:
            if cell[0] == "s":
                get_ss_index(cell[1])

    # Build XML components
    # 1. [Content_Types].xml
    content_types = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
    content_types += '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
    content_types += '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
    content_types += '<Default Extension="xml" ContentType="application/xml"/>'
    content_types += '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'
    content_types += '<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
    content_types += '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>'
    content_types += '<Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>'
    content_types += '</Types>'

    # 2. _rels/.rels
    rels = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
    rels += '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
    rels += '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>'
    rels += '</Relationships>'

    # 3. xl/_rels/workbook.xml.rels
    wb_rels = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
    wb_rels += '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
    wb_rels += '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>'
    wb_rels += '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>'
    wb_rels += '<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/>'
    wb_rels += '</Relationships>'

    # 4. xl/workbook.xml
    workbook = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
    workbook += '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
    workbook += '<sheets>'
    workbook += f'<sheet name="{SHEET_NAME}" sheetId="1" r:id="rId1"/>'
    workbook += '</sheets>'
    workbook += '</workbook>'

    # 5. xl/styles.xml
    # Styles:
    #   numFmt 164: #,##0 (thousands separator)
    #   font 0: default (Calibri 11)
    #   font 1: bold white (for header)
    #   font 2: bold (for title)
    #   fill 0: none
    #   fill 1: gray125 (required)
    #   fill 2: blue header (#4472C4)
    #   border 0: no border
    #   border 1: thin border all sides
    #   xf 0: default
    #   xf 1: header style (bold white, blue fill, centered, thin border)
    #   xf 2: data centered with thin border
    #   xf 3: number with thousands separator, centered, thin border
    #   xf 4: title style (bold, centered)
    #   xf 5: note style (left aligned)
    styles = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
    styles += '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'

    # Number formats
    styles += '<numFmts count="1">'
    styles += '<numFmt numFmtId="164" formatCode="#,##0"/>'
    styles += '</numFmts>'

    # Fonts
    styles += '<fonts count="3">'
    # Font 0: default
    styles += '<font><sz val="11"/><name val="Calibri"/></font>'
    # Font 1: bold white for header
    styles += '<font><b/><sz val="11"/><color rgb="FFFFFFFF"/><name val="Calibri"/></font>'
    # Font 2: bold for title
    styles += '<font><b/><sz val="14"/><name val="Calibri"/></font>'
    styles += '</fonts>'

    # Fills
    styles += '<fills count="3">'
    # Fill 0: none (required)
    styles += '<fill><patternFill patternType="none"/></fill>'
    # Fill 1: gray125 (required)
    styles += '<fill><patternFill patternType="gray125"/></fill>'
    # Fill 2: blue header
    styles += '<fill><patternFill patternType="solid"><fgColor rgb="FF4472C4"/></patternFill></fill>'
    styles += '</fills>'

    # Borders
    styles += '<borders count="2">'
    # Border 0: no border
    styles += '<border><left/><right/><top/><bottom/><diagonal/></border>'
    # Border 1: thin all sides
    styles += '<border>'
    styles += '<left style="thin"><color auto="1"/></left>'
    styles += '<right style="thin"><color auto="1"/></right>'
    styles += '<top style="thin"><color auto="1"/></top>'
    styles += '<bottom style="thin"><color auto="1"/></bottom>'
    styles += '<diagonal/>'
    styles += '</border>'
    styles += '</borders>'

    # Cell style xfs (required)
    styles += '<cellStyleXfs count="1">'
    styles += '<xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>'
    styles += '</cellStyleXfs>'

    # Cell xfs
    styles += '<cellXfs count="6">'
    # xf 0: default
    styles += '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>'
    # xf 1: header (bold white, blue fill, centered, thin border)
    styles += '<xf numFmtId="0" fontId="1" fillId="2" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf>'
    # xf 2: data string centered with thin border
    styles += '<xf numFmtId="0" fontId="0" fillId="0" borderId="1" xfId="0" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf>'
    # xf 3: number with thousands separator, centered, thin border
    styles += '<xf numFmtId="164" fontId="0" fillId="0" borderId="1" xfId="0" applyNumberFormat="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf>'
    # xf 4: title style (bold, centered)
    styles += '<xf numFmtId="0" fontId="2" fillId="0" borderId="0" xfId="0" applyFont="1" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf>'
    # xf 5: note style (left aligned)
    styles += '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyAlignment="1"><alignment horizontal="left"/></xf>'
    styles += '</cellXfs>'

    styles += '</styleSheet>'

    # 6. xl/sharedStrings.xml
    ss_xml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
    ss_xml += f'<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="{len(shared_strings)}" uniqueCount="{len(shared_strings)}">'
    for s in shared_strings:
        # Escape XML special characters
        escaped = s.replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;')
        ss_xml += f'<si><t>{escaped}</t></si>'
    ss_xml += '</sst>'

    # 7. xl/worksheets/sheet1.xml
    # Build sheet with column widths and row data
    total_rows = len(rows_data)
    last_col = col_letter(7)  # H
    dimension = f"A1:{last_col}{total_rows}"

    sheet = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
    sheet += '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
    sheet += f'<dimension ref="{dimension}"/>'
    sheet += '<sheetViews><sheetView tabSelected="1" workbookViewId="0"><selection activeCell="A1" sqref="A1"/></sheetView></sheetViews>'
    sheet += '<sheetFormatPr defaultRowHeight="15"/>'

    # Column widths
    sheet += '<cols>'
    sheet += '<col min="1" max="1" width="8" customWidth="1"/>'   # Age
    sheet += '<col min="2" max="2" width="8" customWidth="1"/>'   # Gender
    sheet += '<col min="3" max="8" width="14" customWidth="1"/>'  # Payment terms
    sheet += '</cols>'

    # Sheet data
    sheet += '<sheetData>'
    for row_idx, row in enumerate(rows_data):
        row_num = row_idx + 1
        sheet += f'<row r="{row_num}">'

        for col_idx, cell in enumerate(row):
            cell_ref = f"{col_letter(col_idx)}{row_num}"

            if cell[0] == "empty":
                continue
            elif cell[0] == "s":
                # String type - reference shared strings
                ss_idx = get_ss_index(cell[1])
                # Determine style
                if row_idx == 0:
                    style = 4  # title
                elif row_idx == 2:
                    style = 1  # header
                elif row_idx == total_rows - 1:
                    style = 5  # note
                else:
                    style = 2  # data centered
                sheet += f'<c r="{cell_ref}" t="s" s="{style}"><v>{ss_idx}</v></c>'
            elif cell[0] == "n":
                # Number type
                if row_idx == 2:
                    style = 1  # header
                elif col_idx >= 2:
                    style = 3  # number with format
                else:
                    style = 2  # centered (age column)
                sheet += f'<c r="{cell_ref}" s="{style}"><v>{cell[1]}</v></c>'

        sheet += '</row>'
    sheet += '</sheetData>'

    # Merge cells for title and note (must come after sheetData per OOXML spec)
    note_row = total_rows
    sheet += '<mergeCells count="2">'
    sheet += f'<mergeCell ref="A1:{last_col}1"/>'
    sheet += f'<mergeCell ref="A{note_row}:{last_col}{note_row}"/>'
    sheet += '</mergeCells>'

    sheet += '</worksheet>'

    # Write the xlsx (ZIP) file
    with zipfile.ZipFile(output_path, 'w', zipfile.ZIP_DEFLATED) as zf:
        zf.writestr('[Content_Types].xml', content_types)
        zf.writestr('_rels/.rels', rels)
        zf.writestr('xl/_rels/workbook.xml.rels', wb_rels)
        zf.writestr('xl/workbook.xml', workbook)
        zf.writestr('xl/styles.xml', styles)
        zf.writestr('xl/sharedStrings.xml', ss_xml)
        zf.writestr('xl/worksheets/sheet1.xml', sheet)

    print(f"Generated: {output_path}")
    print(f"  Sheet: {SHEET_NAME}")
    print(f"  Rows: {len(PREMIUM_DATA)} data rows (ages 0-55, male and female)")
    print(f"  Title: {TITLE}")


if __name__ == "__main__":
    script_dir = os.path.dirname(os.path.abspath(__file__))
    output_file = os.path.join(script_dir, "保费费率表.xlsx")
    create_xlsx(output_file)
