# ops/fixtures.py
"""Deterministic synthetic workbooks. No company data. Used by CI and by the Univer gate."""
import xml.etree.ElementTree as ET
import zipfile
import base64
from openpyxl import Workbook
from openpyxl.comments import Comment
from openpyxl.formatting.rule import CellIsRule, FormulaRule, Rule
from openpyxl.styles import Font, PatternFill, Protection
from openpyxl.styles.differential import DifferentialStyle
from openpyxl.worksheet.table import Table, TableStyleInfo
from openpyxl.utils import get_column_letter
from openpyxl.workbook.defined_name import DefinedName

CF_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"

DRAWING_XML = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
 xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
 xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
 <xdr:oneCellAnchor><xdr:from><xdr:col>1</xdr:col><xdr:colOff>95250</xdr:colOff><xdr:row>1</xdr:row><xdr:rowOff>190500</xdr:rowOff></xdr:from><xdr:ext cx="914400" cy="457200"/><xdr:pic><xdr:nvPicPr><xdr:cNvPr id="2" name="Internal picture"/><xdr:cNvPicPr/></xdr:nvPicPr><xdr:blipFill><a:blip r:embed="rId1"/></xdr:blipFill><xdr:spPr/></xdr:pic><xdr:clientData/></xdr:oneCellAnchor>
 <xdr:twoCellAnchor><xdr:from><xdr:col>2</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>2</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from><xdr:to><xdr:col>4</xdr:col><xdr:colOff>47625</xdr:colOff><xdr:row>4</xdr:row><xdr:rowOff>95250</xdr:rowOff></xdr:to><xdr:sp><xdr:nvSpPr><xdr:cNvPr id="3" name="Plan note"/><xdr:cNvSpPr/></xdr:nvSpPr><xdr:spPr/><xdr:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>Review price</a:t></a:r><a:r><a:t> plan</a:t></a:r></a:p><a:p><a:r><a:t>Second line</a:t></a:r></a:p></xdr:txBody></xdr:sp><xdr:clientData/></xdr:twoCellAnchor>
 <xdr:oneCellAnchor><xdr:from><xdr:col>0</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>5</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from><xdr:ext cx="1371600" cy="381000"/><xdr:sp macro="Module1.Run Report"><xdr:nvSpPr><xdr:cNvPr id="4" name="Run report"/><xdr:cNvSpPr/></xdr:nvSpPr><xdr:spPr/><xdr:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>Run report</a:t></a:r></a:p></xdr:txBody></xdr:sp><xdr:clientData/></xdr:oneCellAnchor>
 <xdr:twoCellAnchor><xdr:from><xdr:col>5</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>1</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from><xdr:to><xdr:col>7</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>3</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to><xdr:graphicFrame><xdr:nvGraphicFramePr><xdr:cNvPr id="5" name="Region slicer"/><xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr><xdr:xfrm/><a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/drawing/2010/slicer"><a:slicer name="Region"/></a:graphicData></a:graphic></xdr:graphicFrame><xdr:clientData/></xdr:twoCellAnchor>
 <xdr:oneCellAnchor><xdr:from><xdr:col>8</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>1</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from><xdr:ext cx="914400" cy="457200"/><xdr:pic><xdr:nvPicPr><xdr:cNvPr id="6" name="External blocked"/><xdr:cNvPicPr/></xdr:nvPicPr><xdr:blipFill><a:blip r:embed="rId2"/></xdr:blipFill><xdr:spPr/></xdr:pic><xdr:clientData/></xdr:oneCellAnchor>
 <xdr:oneCellAnchor><xdr:from><xdr:col>9</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>1</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from><xdr:ext cx="914400" cy="457200"/><xdr:pic><xdr:nvPicPr><xdr:cNvPr id="7" name="Traversal blocked"/><xdr:cNvPicPr/></xdr:nvPicPr><xdr:blipFill><a:blip r:embed="rId3"/></xdr:blipFill><xdr:spPr/></xdr:pic><xdr:clientData/></xdr:oneCellAnchor>
 <xdr:oneCellAnchor><xdr:from><xdr:col>10</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>1</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from><xdr:ext cx="914400" cy="457200"/><xdr:pic><xdr:nvPicPr><xdr:cNvPr id="8" name="Wrong relation blocked"/><xdr:cNvPicPr/></xdr:nvPicPr><xdr:blipFill><a:blip r:embed="rId4"/></xdr:blipFill><xdr:spPr/></xdr:pic><xdr:clientData/></xdr:oneCellAnchor>
</xdr:wsDr>'''

DRAWING_RELS = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
 <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image1.png"/>
 <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="https://example.test/tracker.png" TargetMode="External"/>
 <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../../../outside.png"/>
 <Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="../media/image1.png"/>
</Relationships>'''

TINY_PNG = base64.b64decode(
    "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNk+A8AAQUBAScY42YAAAAASUVORK5CYII="
)

NOTE_TEXT = "Budget <script>alert('x')</script> & literal"


def make_notes(path):
    """Create two worksheets with separate rich-text comment parts."""
    book = Workbook()
    first = book.active
    first.title = "Notes"
    first["A1"] = "Synthetic"
    first.merge_cells("B2:C3")
    first.freeze_panes = "B2"
    first.column_dimensions["B"].width = 18
    first.row_dimensions[2].height = 30
    second = book.create_sheet("Second")
    second["A1"] = "Synthetic"
    book.save(path)
    with zipfile.ZipFile(path, "r") as archive:
        members = {name: archive.read(name) for name in archive.namelist()}
    members["[Content_Types].xml"] = _note_content_types(members["[Content_Types].xml"])
    for sheet in (1, 2):
        members[f"xl/worksheets/_rels/sheet{sheet}.xml.rels"] = _note_sheet_rels(sheet)
    members["xl/comments1.xml"] = _comment_part_one()
    members["xl/comments2.xml"] = _comment_part_two()
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as archive:
        for name, data in members.items():
            archive.writestr(name, data)


def _note_content_types(data):
    root = ET.fromstring(data)
    namespace = "{http://schemas.openxmlformats.org/package/2006/content-types}"
    content_type = "application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml"
    for number in (1, 2):
        root.append(ET.Element(f"{namespace}Override", {
            "PartName": f"/xl/comments{number}.xml", "ContentType": content_type,
        }))
    return ET.tostring(root, encoding="utf-8", xml_declaration=True)


def _note_sheet_rels(number):
    return f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
 <Relationship Id="rIdNote" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments" Target="../comments{number}.xml"/>
</Relationships>'''.encode()


def _comment_part_one():
    return b'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<comments xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
 <authors><author>A &lt;Admin&gt;</author></authors><commentList>
  <comment ref="B2" authorId="0"><text><r><t>Budget &lt;script&gt;</t></r><r><t>alert('x')&lt;/script&gt; &amp; literal</t></r></text></comment>
  <comment ref="B300" authorId="0"><text><t>Far block</t></text></comment>
 </commentList>
</comments>'''


def _comment_part_two():
    return b'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<comments xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
 <authors><author>Reviewer</author></authors><commentList>
  <comment ref="D4" authorId="0"><text><r><t>Second line</t></r><r><t xml:space="preserve"> continues</t></r></text></comment>
 </commentList>
</comments>'''


def make_drawings(path, image_bytes=TINY_PNG):
    """Create a small workbook with pictures, shapes, a macro, and a slicer frame."""
    book = Workbook()
    sheet = book.active
    sheet.title = "Drawings"
    sheet.freeze_panes = "B2"
    sheet.column_dimensions["A"].width = 12
    sheet.column_dimensions["B"].width = 20
    sheet.row_dimensions[2].height = 30
    sheet["A1"] = "Synthetic"
    book.save(path)
    with zipfile.ZipFile(path, "r") as archive:
        members = {name: archive.read(name) for name in archive.namelist()}
    members["xl/worksheets/sheet1.xml"] = _drawing_sheet(
        members["xl/worksheets/sheet1.xml"]
    )
    members["xl/worksheets/_rels/sheet1.xml.rels"] = _sheet_drawing_rels()
    members["xl/drawings/drawing1.xml"] = DRAWING_XML.encode()
    members["xl/drawings/_rels/drawing1.xml.rels"] = DRAWING_RELS.encode()
    members["xl/media/image1.png"] = image_bytes
    members["outside.png"] = TINY_PNG
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as archive:
        for name, data in members.items():
            archive.writestr(name, data)


def _drawing_sheet(data):
    root = ET.fromstring(data)
    root.append(ET.Element(f"{{{CF_NS}}}drawing", {
        "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id": "rIdDrawing",
    }))
    return ET.tostring(root, encoding="utf-8", xml_declaration=True)


def _sheet_drawing_rels():
    return b'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
 <Relationship Id="rIdDrawing" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/drawing1.xml"/>
</Relationships>'''

def make_semantic(path):
    wb = Workbook()
    inp = wb.active
    inp.title = 'Inputs'
    inp.protection.sheet = True
    rows = [('Cost', 10), ('Retail', 20), ('Blank', None), ('EmptyText', ''), ('Rate', '1.234,5')]
    inputs = []
    for i, (label, val) in enumerate(rows, start=2):
        inp.cell(i, 1, label)
        c = inp.cell(i, 2, val)
        c.protection = Protection(locked=False)
        inputs.append(f'Inputs!B{i}')
    lk = wb.create_sheet('Lookup')
    for i in range(1, 11):
        lk.cell(i, 1, f'K{i}')
        lk.cell(i, 2, i * 1.5)
        lk.cell(i, 3, 'A' if i % 2 else 'B')
    calc = wb.create_sheet('Calc')
    formulas = {
        'B2': '=Inputs!B2*2',
        'B3': '=IFERROR(1-Inputs!B2/Inputs!B3,"")',
        'B4': '=IF(Inputs!B4="","blank","set")',
        'B5': '=IF(ISBLANK(Inputs!B4),1,0)',
        'B6': '=IFERROR(1/0,"err")',
        'B7': '=_xlfn.XLOOKUP("K3",Lookup!A1:A10,Lookup!B1:B10,"none")',
        'B8': '=_xlfn.LET(_xlpm.x,Inputs!B2,_xlpm.x+1)',
        'B9': '=_xlfn.SWITCH(Inputs!B2,10,"ten",20,"twenty","other")',
        'B10': '=_xlfn.IFS(Inputs!B2>15,"hi",TRUE,"lo")',
        'B11': '=_xlfn.NUMBERVALUE(Inputs!B6,",",".")',
        'B12': '=SUMIFS(Lookup!B1:B10,Lookup!C1:C10,"A")',
        'B13': '=_xlfn.SUM(_xlfn._xlws.FILTER(Lookup!B1:B10,Lookup!C1:C10="B"))',
        'B14': '=ROWS(_xlfn.UNIQUE(Lookup!C1:C10))',
        'B15': '=SUM(_xlfn.SEQUENCE(3))',
        'B16': '=VLOOKUP("K5",Lookup!A1:C10,2,FALSE)',
        'B17': '=SUBSTITUTE(TRIM("  a  b "),"a","x")',
    }
    for ref, f in formulas.items():
        calc[ref] = f
    wb.save(path)
    return {'inputs': inputs, 'formulas': {f'Calc!{k}': v for k, v in formulas.items()}}

def make_scale(path, rows=4088, cols=312, lookup_rows=100000):
    wb = Workbook(write_only=True)
    lk = wb.create_sheet('Lookup')
    for r in range(1, lookup_rows + 1):
        lk.append([f'K{r}', r * 0.5, f'G{r % 97}'] + [r % 13] * 14)
    grid = wb.create_sheet('Grid')
    for r in range(1, rows + 1):
        row = [f'K{(r * 7) % lookup_rows + 1}', r % 50 + 1]
        for c in range(3, cols + 1):
            row.append(f'=IFERROR(VLOOKUP($A{r},Lookup!$A:$C,2,FALSE)*$B{r}+{c},0)' if c % 3 else f'=SUBSTITUTE(TRIM($A{r}),"K","")')
        grid.append(row)
    wb.save(path)


def make_conditional_formatting(path):
    """Create a synthetic workbook, then inject the tested OOXML CF rules."""
    book = Workbook()
    inputs = book.active
    inputs.title = "Inputs"
    inputs.protection.sheet = True
    values = {
        "A1": 10, "B2": 12, "B3": 4, "C2": "match literal",
        "D2": "duplicate", "D1026": "duplicate", "E2": "prefix-value",
        "G2": "ERR", "H2": 3, "I2": 1, "M2": 1, "P2": "Case",
        "P1026": "case", "Q2": 1, "Q1026": "=1.0", "R2": 1,
        "S2": 1, "T2": 1, "U2": 1, "V257": 0, "W2": 1,
        "AB2": 0, "AB5": 1, "AC6": 0, "AD2": 1, "AE2": 1, "AF2": 1,
    }
    for address, value in values.items():
        inputs[address] = value
    inputs["B2"].protection = Protection(locked=False)
    for address in ("F2", "L2", "N3", "O3"):
        inputs[address] = "#DIV/0!"
        inputs[address].data_type = "e"
    inputs["F2"].font = Font(bold=True)
    inputs["F2"].fill = PatternFill("solid", fgColor="FF112233")
    inputs["AE2"].fill = PatternFill("solid", fgColor="FF445566")
    for address, value in (("J2", "=literal"), ("K2", "#DIV/0!")):
        inputs[address] = value
        inputs[address].data_type = "s"
    calc = book.create_sheet("Calc")
    calc["B2"] = "=Inputs!B2*2"
    calc["B3"] = "=Inputs!B3*2"
    calc["C2"] = "target"
    quoted = book.create_sheet("O'Brien")
    quoted["A1"] = 7
    quoted["B1"] = "target"
    scratch_row = book.create_sheet("Scratch Row")
    scratch_row["J1"], scratch_row["A2"] = "x", "target"
    scratch_column = book.create_sheet("Scratch Column")
    scratch_column["AF1"], scratch_column["A2"] = 1, "target"
    book.create_sheet("Plain")
    book.defined_names.add(DefinedName("GlobalThreshold", attr_text="'Inputs'!$A$1"))
    book.defined_names.add(DefinedName("SimpleLimit", attr_text="10"))
    book.defined_names.add(DefinedName("Threshold", attr_text="99"))
    inputs.defined_names.add(DefinedName("LocalThreshold", attr_text="'Inputs'!$A$1"))
    inputs.defined_names.add(DefinedName("Threshold", attr_text="'Inputs'!$A$1"))
    calc.defined_names.add(DefinedName("Threshold", attr_text="'Calc'!$B$2"))
    book.save(path)
    _inject_conditional_formats(path)


def make_broad_conditional_formatting(path):
    """Create a sparse normal-width sheet with a full-block literal CF rule."""
    book = Workbook()
    sheet = book.active
    sheet.title = "Broad"
    sheet.cell(1, 1, 1)
    sheet.cell(256, 312, "extent")
    sheet.conditional_formatting.add(
        f"A1:{get_column_letter(312)}256",
        FormulaRule(
            formula=["TRUE()"],
            fill=PatternFill(patternType="solid", fgColor="112233"),
        ),
    )
    book.save(path)


def make_volatile_conditional_formatting(path):
    """Create volatile and quoted-function CF expressions over two targets."""
    book = Workbook()
    sheet = book.active
    sheet.title = "Volatile"
    sheet["A1"], sheet["A2"] = 1, 1
    sheet["B1"], sheet["B2"] = "RAND()", "RAND()"
    sheet["C1"], sheet["C2"] = 1, 1
    sheet["D1"], sheet["D2"] = "VolatileName", "VolatileName"
    book.defined_names.add(DefinedName("VolatileName", attr_text="RAND()"))
    sheet.conditional_formatting.add(
        "A1:A2",
        FormulaRule(
            formula=["RAND()>0.5"],
            fill=PatternFill(patternType="solid", fgColor="112233"),
        ),
    )
    sheet.conditional_formatting.add(
        "B1:B2",
        FormulaRule(
            formula=['EXACT("RAND()","RAND()")'],
            font=Font(bold=True),
        ),
    )
    sheet.conditional_formatting.add(
        "C1:C2",
        FormulaRule(
            formula=["VolatileName>0.5"],
            fill=PatternFill(patternType="solid", fgColor="445566"),
        ),
    )
    sheet.conditional_formatting.add(
        "D1:D2",
        FormulaRule(
            formula=['"VolatileName"="VolatileName"'],
            font=Font(bold=True),
        ),
    )
    book.save(path)


def make_partial_error_conditional_formatting(path):
    """Create one CF rule whose formula errors on only its first target."""
    book = Workbook()
    sheet = book.active
    sheet.title = "Partial"
    for row, value in enumerate((-1, 4, 9), start=1):
        sheet.cell(row, 1, value)
        sheet.cell(row, 2, "target")
    sheet.conditional_formatting.add(
        "B1:B3",
        FormulaRule(
            formula=["SQRT(A1)>1"],
            fill=PatternFill(patternType="solid", fgColor="112233"),
        ),
    )
    book.save(path)


def make_tables(path):
    """Create two sparse tables, an overlapping sheet filter, and custom DXFs."""
    book = Workbook()
    sheet = book.active
    sheet.title = "Tables"
    for cell, value in {
        "A1": "Region", "B1": "Price", "C1": "Status", "D1": "Sheet only",
        "A2": "East", "B2": 10, "C2": "Keep", "D2": "Open",
        "A3": "West", "B3": 20, "C3": "Drop", "D3": "Closed",
        "A256": "East", "B256": 30, "C256": "Keep",
        "A257": "West", "B257": 40, "C257": "Drop",
        "A260": "North", "B260": 50, "C260": "Keep",
        "E1": "Item", "F1": "Amount", "G1": "State",
        "E2": "One", "F2": 11, "G2": "Open",
        "E3": "Two", "F3": 22, "G3": "Closed",
        "E4": "Three", "F4": 33, "G4": "Open",
    }.items():
        sheet[cell] = value
    sheet["F2"].protection = Protection(locked=False)
    sheet.protection.sheet = True
    sheet.auto_filter.ref = "A1:D260"
    built_in = Table(displayName="Orders", ref="A1:C260")
    built_in.tableStyleInfo = TableStyleInfo(
        name="TableStyleMedium2",
        showFirstColumn=True,
        showLastColumn=False,
        showRowStripes=True,
        showColumnStripes=False,
    )
    custom = Table(displayName="CustomPlan", ref="E1:G4")
    custom.tableStyleInfo = TableStyleInfo(
        name="CustomPlanning",
        showFirstColumn=False,
        showLastColumn=False,
        showRowStripes=True,
        showColumnStripes=False,
    )
    sheet.add_table(built_in)
    sheet.add_table(custom)
    sheet.conditional_formatting.add(
        "F2",
        FormulaRule(
            formula=["TRUE()"],
            fill=PatternFill(patternType="solid", fgColor="FF0000"),
        ),
    )
    book.save(path)
    _inject_custom_table_style(path)


def _inject_custom_table_style(path):
    with zipfile.ZipFile(path, "r") as archive:
        members = {name: archive.read(name) for name in archive.namelist()}
    root = ET.fromstring(members["xl/styles.xml"])
    dxfs = root.find(f"{{{CF_NS}}}dxfs")
    if dxfs is None:
        dxfs = ET.SubElement(root, f"{{{CF_NS}}}dxfs")
    base = len(dxfs)
    _append_table_dxf(dxfs, "whole")
    _append_table_dxf(dxfs, "header")
    _append_table_dxf(dxfs, "stripe")
    dxfs.set("count", str(len(dxfs)))
    styles = root.find(f"{{{CF_NS}}}tableStyles")
    custom = ET.SubElement(styles, f"{{{CF_NS}}}tableStyle", {
        "name": "CustomPlanning", "pivot": "0", "table": "1", "count": "3",
    })
    for kind, offset in (("wholeTable", 0), ("headerRow", 1), ("firstRowStripe", 2)):
        ET.SubElement(custom, f"{{{CF_NS}}}tableStyleElement", {
            "type": kind, "dxfId": str(base + offset),
        })
    styles.set("count", str(int(styles.get("count", "0")) + 1))
    members["xl/styles.xml"] = ET.tostring(root, encoding="utf-8", xml_declaration=True)
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as archive:
        for name, data in members.items():
            archive.writestr(name, data)


def _append_table_dxf(dxfs, kind):
    dxf = ET.SubElement(dxfs, f"{{{CF_NS}}}dxf")
    if kind == "whole":
        font = ET.SubElement(dxf, f"{{{CF_NS}}}font")
        ET.SubElement(font, f"{{{CF_NS}}}color", {"rgb": "FF333333"})
        ET.SubElement(dxf, f"{{{CF_NS}}}numFmt", {
            "numFmtId": "176", "formatCode": "$0.00",
        })
        border = ET.SubElement(dxf, f"{{{CF_NS}}}border")
        for side in ("left", "right", "top", "bottom"):
            node = ET.SubElement(border, f"{{{CF_NS}}}{side}", {"style": "thin"})
            ET.SubElement(node, f"{{{CF_NS}}}color", {"rgb": "FF112233"})
        for side, style, color in (("vertical", "dotted", "FF445566"),
                                   ("horizontal", "dashed", "FF778899")):
            node = ET.SubElement(border, f"{{{CF_NS}}}{side}", {"style": style})
            ET.SubElement(node, f"{{{CF_NS}}}color", {"rgb": color})
        return
    if kind == "header":
        ET.SubElement(ET.SubElement(dxf, f"{{{CF_NS}}}font"), f"{{{CF_NS}}}b")
        color = "FFABCDEF"
    else:
        color = "FFEEDDCC"
    pattern = ET.SubElement(ET.SubElement(dxf, f"{{{CF_NS}}}fill"),
                            f"{{{CF_NS}}}patternFill", {"patternType": "solid"})
    ET.SubElement(pattern, f"{{{CF_NS}}}fgColor", {"rgb": color})
    ET.SubElement(pattern, f"{{{CF_NS}}}bgColor", {"rgb": color})


def _inject_conditional_formats(path):
    with zipfile.ZipFile(path, "r") as archive:
        members = {name: archive.read(name) for name in archive.namelist()}
    members["xl/styles.xml"] = _with_dxfs(members["xl/styles.xml"])
    members["xl/worksheets/sheet1.xml"] = _with_rules(
        members["xl/worksheets/sheet1.xml"],
    )
    extra_rules = {
        2: (("C2", 'type="expression" dxfId="0" priority="1"><formula>Threshold=24</formula>'),),
        3: (("B1", 'type="expression" dxfId="2" priority="1"><formula>Threshold=99</formula>'),),
        4: (("A2", 'type="expression" dxfId="0" priority="1"><formula>COUNTIF($1:$1,&quot;x&quot;)=1</formula>'),),
        5: (("A2", 'type="expression" dxfId="1" priority="1"><formula>COUNT(A:AF)=1</formula>'),),
    }
    for sheet, rules in extra_rules.items():
        name = f"xl/worksheets/sheet{sheet}.xml"
        members[name] = _with_rules(members[name], rules)
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as archive:
        for name, data in members.items():
            archive.writestr(name, data)


def _with_dxfs(data):
    root = ET.fromstring(data)
    styles = (
        '<fill><patternFill patternType="solid"><fgColor rgb="FF00FFFF"/><bgColor rgb="FFFF0000"/></patternFill></fill>',
        '<font><color rgb="FF00AA00"/></font>',
        '<font><b val="1"/></font>',
        '<border><bottom style="thin"><color rgb="FF0000FF"/></bottom></border>',
        '<numFmt numFmtId="164" formatCode="0.000"/>',
        '<font><b val="0"/></font><fill><patternFill patternType="none"/></fill>',
        '<fill><patternFill patternType="solid"><fgColor rgb="FF111111"/></patternFill></fill>',
        '<font><b val="1"/></font>',
        '<font><color theme="1"/></font>',
        '<fill><patternFill patternType="solid"><fgColor indexed="10"/></patternFill></fill>',
        '<border><top style="thin"><color rgb="FF000000" tint="0.5"/></top></border>',
        '<fill><patternFill patternType="solid"/></fill>',
    )
    dxfs = ET.Element(f"{{{CF_NS}}}dxfs", {"count": str(len(styles))})
    for style in styles:
        dxfs.append(ET.fromstring(f'<dxf xmlns="{CF_NS}">{style}</dxf>'))
    root.append(dxfs)
    return ET.tostring(root, encoding="utf-8", xml_declaration=True)


def _with_rules(data, rules=None):
    root = ET.fromstring(data)
    rules = rules or (
        ("A2:A3", 'type="expression" dxfId="0" priority="1"><formula>AND(Calc!B2&gt;$A$1,MOD(B2,2)=0)</formula>'),
        ("B2", 'type="cellIs" operator="greaterThan" dxfId="1" priority="2"><formula>10</formula>'),
        ("C2", 'type="containsText" operator="containsText" text="match" dxfId="2" priority="3"><formula>NOT(ISERROR(SEARCH(&quot;match&quot;,C2)))</formula>'),
        ("D2:D1026", 'type="duplicateValues" dxfId="3" priority="4">'),
        ("E2", 'type="beginsWith" operator="beginsWith" text="prefix" dxfId="4" priority="5"><formula>LEFT(E2,LEN(&quot;prefix&quot;))=&quot;prefix&quot;</formula>'),
        ("F2:G2", 'type="containsErrors" dxfId="5" priority="6"><formula>ISERROR(F2)</formula>'),
        ("H2", 'type="cellIs" operator="greaterThan" dxfId="6" priority="7" stopIfTrue="1"><formula>1</formula>'),
        ("H2", 'type="cellIs" operator="greaterThan" dxfId="7" priority="8"><formula>0</formula>'),
        ("I2", 'type="expression" dxfId="0" priority="9"><formula>AND(ROW()=2,COLUMN()=9)</formula>'),
        ("J2", 'type="expression" dxfId="1" priority="10"><formula>EXACT(J2,&quot;=literal&quot;)</formula>'),
        ("K2", 'type="expression" dxfId="2" priority="11"><formula>EXACT(K2,&quot;#DIV/0!&quot;)</formula>'),
        ("L2", 'type="expression" dxfId="3" priority="12"><formula>ISERROR(L2)</formula>'),
        ("M2", 'type="expression" dxfId="4" priority="13"><formula>AND(GlobalThreshold=10,LocalThreshold=10,SimpleLimit=10,Threshold=10)</formula>'),
        ("N2", 'type="containsText" text="none" dxfId="0" priority="14">'),
        ("N3", 'type="containsText" text="div" dxfId="0" priority="15">'),
        ("O2", 'type="beginsWith" text="none" dxfId="0" priority="16">'),
        ("O3", 'type="beginsWith" text="{" dxfId="0" priority="17">'),
        ("P2:P1026", 'type="duplicateValues" dxfId="0" priority="18">'),
        ("Q2:Q1026", 'type="duplicateValues" dxfId="1" priority="19">'),
        ("R2", 'type="expression" dxfId="2" priority="20"><formula>COUNTIF(P:P,&quot;case&quot;)=2</formula>'),
        ("S2", 'type="cellIs" operator="greaterThan" dxfId="8" priority="21"><formula>0</formula>'),
        ("T2", 'type="cellIs" operator="greaterThan" dxfId="9" priority="22"><formula>0</formula>'),
        ("U2", 'type="cellIs" operator="greaterThan" dxfId="10" priority="23"><formula>0</formula>'),
        ("V2:V257", 'type="expression" dxfId="0" priority="24"><formula>ROW()&gt;0</formula>'),
        ("V2:V257", 'type="expression" dxfId="1" priority="25"><formula>COLUMN()=22</formula>'),
        ("W2", 'type="expression" dxfId="2" priority="26"><formula>COUNTIF(X:X,5)=1</formula>'),
        ("AA2:AA3 AC5:AC6", 'type="expression" dxfId="0" priority="27"><formula>$AB2&gt;0</formula>'),
        ("AD2", 'type="expression" dxfId="2" priority="28"><formula>\'O\'\'Brien\'!A1=7</formula>'),
        ("AE2", 'type="cellIs" operator="greaterThan" dxfId="11" priority="29"><formula>0</formula>'),
        ("XFD:XFD", 'type="expression" dxfId="0" priority="30"><formula>TRUE</formula>'),
        ("AF2000:AF2001", 'type="expression" dxfId="0" priority="31"><formula>TRUE</formula>'),
    )
    for sqref, body in rules:
        conditional = ET.Element(f"{{{CF_NS}}}conditionalFormatting", {"sqref": sqref})
        conditional.append(ET.fromstring(f'<cfRule xmlns="{CF_NS}" {body}</cfRule>'))
        root.append(conditional)
    return ET.tostring(root, encoding="utf-8", xml_declaration=True)


def make_fidelity(path):
    """Create one synthetic workbook carrying every fidelity feature at once."""
    book = Workbook()
    rules = book.active
    rules.title = "CF Rules"
    for address, value in {
        "A1": "Target", "B1": "Threshold", "C1": "Text", "D1": "Duplicate",
        "E1": "Prefix", "F1": "Broken", "G1": "Plan",
        "A2": 10, "A3": 2, "B2": 10, "B3": 1, "C2": "match literal",
        "D2": "dup", "D3": "dup", "E2": "prefix-value", "F2": "#DIV/0!",
        "G2": 4,
    }.items():
        rules[address] = value
    rules["F2"].data_type = "e"
    rules["A4"] = 0.25
    rules["A4"].number_format = "0.00"
    rules.conditional_formatting.add(
        "A2:A3",
        FormulaRule(formula=["$B2>5"],
                    fill=PatternFill(patternType="solid", fgColor="FFFFC7CE")),
    )
    rules.conditional_formatting.add(
        "B2:B3",
        CellIsRule(operator="greaterThan", formula=["5"],
                   fill=PatternFill(patternType="solid", fgColor="FFD9EAD3")),
    )
    for kind, area, text, fill in (
        ("containsText", "C2:C3", "match", "FF9C6500"),
        ("duplicateValues", "D2:D3", None, "FFEBF1DE"),
        ("beginsWith", "E2:E3", "prefix", "FFCCFFFF"),
    ):
        extra = {"text": text} if text else {}
        operator = kind if kind in {"containsText", "beginsWith"} else None
        rules.conditional_formatting.add(
            area,
            Rule(
                type=kind,
                operator=operator,
                dxf=DifferentialStyle(
                    fill=PatternFill(patternType="solid", fgColor=fill),
                ),
                **extra,
            ),
        )
    rules.conditional_formatting.add(
        "F2:F3",
        Rule(
            type="containsErrors",
            dxf=DifferentialStyle(font=Font(bold=True, color="FF9C0006")),
        ),
    )

    drawings = book.create_sheet("Drawings")
    drawings["A1"] = "Synthetic drawing sheet"
    drawings["A3"] = "Picture, text box, macro and slicer sit on this sheet"

    notes_one = book.create_sheet("Notes One")
    notes_one["A1"] = "Hover the red marker"
    notes_one["B2"] = "Annotated cell"
    notes_one["B2"].comment = Comment(
        "Synthetic note: confirm the price plan before release.", "Planner",
    )
    notes_two = book.create_sheet("Notes Two")
    notes_two["A1"] = "Second note sheet"
    notes_two["D4"] = "Second annotated cell"
    notes_two["D4"].comment = Comment("Synthetic note on the second sheet.", "Reviewer")

    tables = book.create_sheet("Tables")
    for cell, value in {
        "A1": "Region", "B1": "Price", "C1": "Status", "D1": "Sheet only",
        "A2": "East", "B2": 10, "C2": "Keep", "D2": "Open",
        "A3": "West", "B3": 20, "C3": "Drop", "D3": "Closed",
        "A4": "North", "B4": 30, "C4": "Keep",
        "F1": "Item", "G1": "Amount",
        "F2": "One", "G2": 11, "F3": "Two", "G3": 22, "F4": "Three", "G4": 33,
    }.items():
        tables[cell] = value
    tables.auto_filter.ref = "A1:D4"
    for name, ref, style in (
        ("Orders", "A1:C4", "TableStyleMedium2"),
        ("PlanItems", "F1:G4", "TableStyleLight9"),
    ):
        table = Table(displayName=name, ref=ref)
        table.tableStyleInfo = TableStyleInfo(
            name=style, showFirstColumn=False, showLastColumn=False,
            showRowStripes=True, showColumnStripes=False,
        )
        tables.add_table(table)

    view = book.create_sheet("View Demo")
    view.sheet_properties.tabColor = "FF107C41"
    view.sheet_view.showGridLines = False
    view.sheet_view.zoomScale = 125
    view.freeze_panes = "D6"
    view.column_dimensions["B"].width = 18
    view.row_dimensions[2].height = 30
    view.row_dimensions[5].hidden = True
    view.column_dimensions["D"].hidden = True
    view.merge_cells("B6:C7")
    view["A1"] = "Sheet view"
    view["B1"] = "Input"
    view["B2"] = 24
    view["B2"].protection = Protection(locked=False)
    view["C2"] = "=B2*2"
    view["B6"] = "Merged header"
    view.protection.sheet = True
    book.create_sheet("Hidden Demo").sheet_state = "hidden"
    book.create_sheet("Very Hidden Demo").sheet_state = "veryHidden"
    book.active = book.sheetnames.index("View Demo")
    book.save(path)
    _attach_fidelity_parts(path)


def _attach_fidelity_parts(path):
    """Add the drawing part and the frozen pane that openpyxl cannot express."""
    with zipfile.ZipFile(path, "r") as archive:
        members = {name: archive.read(name) for name in archive.namelist()}
    members["xl/worksheets/sheet2.xml"] = _drawing_sheet(
        members["xl/worksheets/sheet2.xml"]
    )
    members["xl/worksheets/_rels/sheet2.xml.rels"] = _sheet_drawing_rels()
    members["xl/drawings/drawing1.xml"] = DRAWING_XML.encode()
    members["xl/drawings/_rels/drawing1.xml.rels"] = DRAWING_RELS.encode()
    members["xl/media/image1.png"] = TINY_PNG
    members["xl/worksheets/sheet6.xml"] = _frozen_pane(
        members["xl/worksheets/sheet6.xml"], 1, 2, "D6",
    )
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as archive:
        for name, data in members.items():
            archive.writestr(name, data)


def _frozen_pane(data, columns, rows, top_left):
    root = ET.fromstring(data)
    for pane in root.iter(f"{{{CF_NS}}}pane"):
        pane.set("xSplit", str(columns))
        pane.set("ySplit", str(rows))
        pane.set("topLeftCell", top_left)
    return ET.tostring(root, encoding="utf-8", xml_declaration=True)
