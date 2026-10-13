"""Sheet pictures and shapes import where Excel draws them; hidden, macro and fallback objects stay out."""

import base64
import io
import unittest
import xml.etree.ElementTree as ET
import zipfile

XDR = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
A = "http://schemas.openxmlformats.org/drawingml/2006/main"
R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
MC = "http://schemas.openxmlformats.org/markup-compatibility/2006"
PNG = base64.b64decode("iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR4nGNgYGD4DwABBAEAwS2OUAAAAABJRU5ErkJggg==")


def anchor(body, start=(1, 2), end=(4, 5)):
    point = "<xdr:{0}><xdr:col>{2}</xdr:col><xdr:colOff>9525</xdr:colOff><xdr:row>{1}</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:{0}>"
    return (f"<xdr:twoCellAnchor>{point.format('from', *start)}{point.format('to', *end)}{body}"
            "<xdr:clientData/></xdr:twoCellAnchor>")


def shape(name, *, hidden=False, macro="", fill="", text="", link=""):
    hide = ' hidden="1"' if hidden else ""
    click = f'<a:hlinkClick r:id="{link}"/>' if link else ""
    return (f'<xdr:sp macro="{macro}" textlink=""><xdr:nvSpPr><xdr:cNvPr id="2" name="{name}"{hide}>{click}</xdr:cNvPr>'
            '<xdr:cNvSpPr/></xdr:nvSpPr><xdr:spPr><a:prstGeom prst="roundRect"><a:avLst><a:gd name="adj" fmla="val 25000"/>'
            f'</a:avLst></a:prstGeom>{fill}<a:effectLst><a:outerShdw blurRad="40000"/></a:effectLst></xdr:spPr>'
            '<xdr:style><a:lnRef idx="2"><a:schemeClr val="accent1"><a:shade val="50000"/></a:schemeClr></a:lnRef>'
            '<a:fillRef idx="1"><a:schemeClr val="accent1"/></a:fillRef><a:effectRef idx="0"><a:schemeClr val="accent1"/>'
            '</a:effectRef><a:fontRef idx="minor"><a:schemeClr val="lt1"/></a:fontRef></xdr:style>'
            '<xdr:txBody><a:bodyPr anchor="ctr" lIns="0"/><a:lstStyle/><a:p><a:pPr algn="ctr"/>'
            f'<a:r><a:rPr lang="en-US" sz="1200" b="1"/><a:t>{text}</a:t></a:r></a:p></xdr:txBody></xdr:sp>')


DRAWING = (
    f'<xdr:wsDr xmlns:xdr="{XDR}" xmlns:a="{A}" xmlns:r="{R}" xmlns:mc="{MC}">'
    + anchor('<xdr:pic><xdr:nvPicPr><xdr:cNvPr id="1" name="Logo"/><xdr:cNvPicPr/></xdr:nvPicPr>'
             '<xdr:blipFill><a:blip r:embed="rImg"/></xdr:blipFill><xdr:spPr><a:prstGeom prst="rect"/></xdr:spPr></xdr:pic>')
    + anchor(shape("Nav", text="Go", link="rLink"), (0, 0), (1, 1))
    + anchor(shape("Plain", fill='<a:solidFill><a:srgbClr val="00B050"/></a:solidFill>', text="x"))
    + anchor(shape("Hidden", hidden=True, text="h"))
    + anchor(shape("Button", macro="[0]!Run", text="b"))
    + anchor(shape("Empty", fill="<a:noFill/>").replace('<a:lnRef idx="2">', '<a:lnRef idx="0">'))
    + '<mc:AlternateContent><mc:Choice Requires="a14">'
    + anchor('<xdr:graphicFrame macro=""><xdr:nvGraphicFramePr><xdr:cNvPr id="9" name="Slicer"/>'
             '<xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr></xdr:graphicFrame>')
    + "</mc:Choice><mc:Fallback>" + anchor(shape("Fallback", text="old Excel")) + "</mc:Fallback></mc:AlternateContent>"
    + "</xdr:wsDr>")
RELS = ('<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rImg" Type="x/image" Target="../media/image1.png"/>'
        '<Relationship Id="rLink" Type="x/hyperlink" Target="#\'Plan \'\'A\'\'\'!$B$7"/></Relationships>')


class DrawingImportTests(unittest.TestCase):
    def read(self):
        from pp.template.colors import Palette
        from pp.template.drawings import drawing_objects

        buffer = io.BytesIO()
        with zipfile.ZipFile(buffer, "w") as archive:
            archive.writestr("xl/drawings/drawing1.xml", DRAWING)
            archive.writestr("xl/drawings/_rels/drawing1.xml.rels", RELS)
            archive.writestr("xl/media/image1.png", PNG)
        with zipfile.ZipFile(buffer) as archive:
            return drawing_objects(ET.fromstring(DRAWING), "xl/drawings/drawing1.xml", archive, Palette())

    def test_visible_pictures_and_shapes_only(self):
        found = self.read()
        self.assertEqual([item["kind"] for item in found], ["picture", "shape", "shape"])
        picture, nav, plain = found
        self.assertEqual(picture["from"], {"row": 1, "col": 2, "dx": 1.0, "dy": 0.0})
        self.assertEqual(picture["to"], {"row": 4, "col": 5, "dx": 1.0, "dy": 0.0})
        self.assertEqual(picture["src"], "data:image/png;base64," + base64.b64encode(PNG).decode())
        self.assertEqual(nav["link"], {"sheet": "Plan 'A'", "cell": "B7"})
        self.assertEqual((nav["geom"], nav["adjust"], nav["shadow"]), ("roundRect", 0.25, True))
        # theme style references: accent1 fill, accent1 at half shade for the line, light text
        self.assertEqual(nav["fill"], "#4472C4")
        self.assertEqual(nav["line"]["width"], 2.67)
        self.assertEqual(plain["fill"], "#00B050")
        self.assertEqual(nav["text"]["anchor"], "ctr")
        self.assertEqual(nav["text"]["inset"][0], 0)
        self.assertEqual(nav["text"]["paragraphs"], [{"align": "ctr", "runs": [
            {"t": "Go", "b": True, "i": False, "sz": 16.0, "color": "#FFFFFF", "font": None}]}])

    def test_macro_shape_buttons_keep_the_shape_look(self):
        from pp.template.colors import Palette
        from pp.template.importer import _drawing_buttons

        buffer = io.BytesIO()
        with zipfile.ZipFile(buffer, "w") as archive:
            archive.writestr("xl/drawings/drawing1.xml", DRAWING)
        with zipfile.ZipFile(buffer) as archive:
            buttons = _drawing_buttons(ET.fromstring(DRAWING), "xl/drawings/drawing1.xml", archive, Palette())
        self.assertEqual([button["macro"] for button in buttons], ["Run"])
        look = buttons[0]["look"]
        self.assertEqual((look["geom"], look["fill"]), ("roundRect", "#4472C4"))
        self.assertEqual(look["text"]["paragraphs"][0]["runs"][0]["t"], "b")
        self.assertNotIn("link", look)

    def test_shape_that_does_not_resize_with_cells_keeps_its_own_size(self):
        from pp.template.drawings import anchor_box

        drawing = ET.fromstring(DRAWING.replace("<xdr:twoCellAnchor>", '<xdr:twoCellAnchor editAs="oneCell">', 1)
                                .replace('<xdr:spPr><a:prstGeom prst="rect"/>',
                                         '<xdr:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="952500" cy="476250"/></a:xfrm>'
                                         '<a:prstGeom prst="rect"/>', 1))
        anchor = next(iter(drawing))
        self.assertEqual(anchor_box(anchor, anchor.find(f"{{{XDR}}}pic")),
                         {"from": {"row": 1, "col": 2, "dx": 1.0, "dy": 0.0}, "size": {"w": 100.0, "h": 50.0}})


if __name__ == "__main__":
    unittest.main()
