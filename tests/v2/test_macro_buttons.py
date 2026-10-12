import json
import tempfile
import unittest
import zipfile
from pathlib import Path

from openpyxl import Workbook

from pp.template.importer import import_template
from pp.workbook_service import TemplateSnapshot


DRAWING_NS = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
MAIN_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"
REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
PKG_RELS = "http://schemas.openxmlformats.org/package/2006/relationships"
DRAWING_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing"
VML_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/vmlDrawing"

DRAWING_XML = f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<xdr:wsDr xmlns:xdr="{DRAWING_NS}" xmlns:a="{MAIN_NS}">
 <xdr:twoCellAnchor>
  <xdr:from><xdr:col>1</xdr:col><xdr:colOff>95250</xdr:colOff><xdr:row>2</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>
  <xdr:to><xdr:col>3</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>4</xdr:row><xdr:rowOff>190500</xdr:rowOff></xdr:to>
  <xdr:sp macro="[0]!Sheet13.resetX">
   <xdr:nvSpPr><xdr:cNvPr id="2" name="Button 1"/><xdr:cNvSpPr/></xdr:nvSpPr>
   <xdr:txBody><a:p><a:r><a:t>Reset </a:t></a:r><a:r><a:t>inputs</a:t></a:r></a:p></xdr:txBody>
  </xdr:sp>
 </xdr:twoCellAnchor>
 <xdr:oneCellAnchor>
  <xdr:from><xdr:col>1</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>6</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>
  <xdr:ext cx="952500" cy="285750"/>
  <xdr:grpSp>
   <xdr:nvGrpSpPr><xdr:cNvPr id="3" name="Group 2"/><xdr:cNvGrpSpPr/></xdr:nvGrpSpPr>
   <xdr:sp macro="[0]!Plain">
    <xdr:nvSpPr><xdr:cNvPr id="4" name="Button 2"/><xdr:cNvSpPr/></xdr:nvSpPr>
   </xdr:sp>
  </xdr:grpSp>
 </xdr:oneCellAnchor>
 <xdr:twoCellAnchor>
  <xdr:from><xdr:col>0</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>9</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>
  <xdr:to><xdr:col>1</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>10</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>
  <xdr:sp macro="[0]!Hidden">
   <xdr:nvSpPr><xdr:cNvPr id="5" name="Hidden 3" hidden="1"/><xdr:cNvSpPr/></xdr:nvSpPr>
  </xdr:sp>
 </xdr:twoCellAnchor>
 <xdr:twoCellAnchor>
  <xdr:from><xdr:col>0</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>11</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>
  <xdr:to><xdr:col>1</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>12</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>
  <xdr:sp>
   <xdr:nvSpPr><xdr:cNvPr id="6" name="Plain shape"/><xdr:cNvSpPr/></xdr:nvSpPr>
  </xdr:sp>
 </xdr:twoCellAnchor>
</xdr:wsDr>
"""

VML_XML = """<xml xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office"
 xmlns:x="urn:schemas-microsoft-com:office:excel">
 <v:shape id="_x0000_s1025" type="#_x0000_t201" style="position:absolute;visibility:visible">
  <v:textbox><div style="text-align:center"><font>Run</font></div></v:textbox>
  <x:ClientData ObjectType="Button">
   <x:Anchor>1, 5, 3, 0, 4, 10, 6, 0</x:Anchor>
   <x:FmlaMacro>'Book.xlsm'!Recalc</x:FmlaMacro>
  </x:ClientData>
 </v:shape>
 <v:shape id="_x0000_s1026" type="#_x0000_t201" style="position:absolute;visibility:hidden">
  <v:textbox><div>Secret</div></v:textbox>
  <x:ClientData ObjectType="Button">
   <x:Anchor>0, 0, 0, 0, 1, 0, 1, 0</x:Anchor>
   <x:FmlaMacro>[0]!Hidden</x:FmlaMacro>
  </x:ClientData>
 </v:shape>
 <v:shape id="_x0000_s1027" type="#_x0000_t201" style="position:absolute;visibility:visible">
  <x:ClientData ObjectType="Checkbox">
   <x:Anchor>0, 0, 0, 0, 1, 0, 1, 0</x:Anchor>
   <x:FmlaMacro>[0]!NotAButton</x:FmlaMacro>
  </x:ClientData>
 </v:shape>
</xml>
"""


def write_button_fixture(path):
    """Synthetic workbook: DrawingML buttons on sheet 0, one VML button on sheet 1, none on sheet 2."""
    workbook = Workbook()
    workbook.active.title = "Inputs"
    workbook.create_sheet("Lookup")
    workbook.create_sheet("Calc")
    workbook.save(path)
    add_buttons(path)


def add_buttons(path):
    """Add the drawing and VML buttons to sheets 1 and 2 of a saved workbook (sheets 0-2 must exist)."""
    with zipfile.ZipFile(path) as source:
        members = {info.filename: source.read(info.filename) for info in source.infolist()}
    sheet_rels = (f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                  f'<Relationships xmlns="{PKG_RELS}">'
                  f'<Relationship Id="rIdD" Type="{DRAWING_REL}" Target="../drawings/drawing1.xml"/>'
                  f'</Relationships>')
    lookup_rels = (f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                   f'<Relationships xmlns="{PKG_RELS}">'
                   f'<Relationship Id="rIdV" Type="{VML_REL}" Target="/xl/drawings/vmlDrawing1.vml"/>'
                   f'</Relationships>')
    members["xl/drawings/drawing1.xml"] = DRAWING_XML.encode("utf-8")
    members["xl/drawings/vmlDrawing1.vml"] = VML_XML.encode("utf-8")
    types = members["[Content_Types].xml"].decode("utf-8").replace(
        "</Types>",
        '<Default Extension="vml" ContentType="application/vnd.openxmlformats-officedocument.vmlDrawing"/>'
        '<Override PartName="/xl/drawings/drawing1.xml" '
        'ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/></Types>')
    members["[Content_Types].xml"] = types.encode("utf-8")
    members["xl/worksheets/_rels/sheet1.xml.rels"] = sheet_rels.encode("utf-8")
    members["xl/worksheets/_rels/sheet2.xml.rels"] = lookup_rels.encode("utf-8")
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as target:
        for name, data in members.items():
            target.writestr(name, data)


def expected_buttons():
    return [
        {"id": "b0", "sheet": 0, "from": {"row": 2, "col": 1, "dx": 10.0, "dy": 0.0},
         "to": {"row": 4, "col": 3, "dx": 0.0, "dy": 20.0}, "label": "Reset inputs",
         "macro": "Sheet13.resetX"},
        {"id": "b1", "sheet": 0, "from": {"row": 6, "col": 1, "dx": 0.0, "dy": 0.0},
         "size": {"w": 100.0, "h": 30.0}, "label": "Button 2", "macro": "Plain"},
        {"id": "b2", "sheet": 1, "from": {"row": 3, "col": 1, "dx": 5, "dy": 0},
         "to": {"row": 6, "col": 4, "dx": 10, "dy": 0}, "label": "Run", "macro": "Recalc"},
    ]


class MacroButtonImportTests(unittest.TestCase):
    def setUp(self):
        self._temporary = tempfile.TemporaryDirectory()
        self.workspace = Path(self._temporary.name)
        self.source = self.workspace / "buttons.xlsm"
        write_button_fixture(self.source)

    def tearDown(self):
        self._temporary.cleanup()

    def test_import_extracts_drawingml_and_vml_buttons_with_offsets_and_names(self):
        snapshot = import_template(self.source, self.workspace / "data")
        buttons = json.loads((snapshot / "buttons.json").read_text(encoding="utf-8"))
        self.assertEqual(buttons, expected_buttons())
        self.assertEqual(TemplateSnapshot(snapshot).buttons, expected_buttons())

    def test_snapshot_without_buttons_file_loads_with_no_buttons(self):
        snapshot = import_template(self.source, self.workspace / "data")
        (snapshot / "buttons.json").unlink()
        self.assertEqual(TemplateSnapshot(snapshot).buttons, [])


if __name__ == "__main__":
    unittest.main()
