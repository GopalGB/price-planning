"""Excel colors resolve to the hex Excel itself shows."""

import io
import unittest
import xml.etree.ElementTree as ET
import zipfile

NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
A = "http://schemas.openxmlformats.org/drawingml/2006/main"


def color(**attrs):
    return ET.fromstring(f'<color xmlns="{NS}" ' + " ".join(f'{k}="{v}"' for k, v in attrs.items()) + "/>")


class CellColorTests(unittest.TestCase):
    def setUp(self):
        from pp.template.colors import Palette

        self.palette = Palette()

    def test_theme_tints_match_excel_swatches(self):
        # Excel's swatch hex for theme + tint pairs (Office theme); Excel rounds on its own HLS grid,
        # so a channel may differ by up to 2 of 255, which is not visible
        cases = [(4, 0.7999816888943144, "#D9E1F2"), (4, 0.5999938962981048, "#B4C6E7"),
                 (4, 0.3999755851924192, "#8EA9DB"), (4, -0.249977111117893, "#2F5597"),
                 (4, -0.499984740745262, "#203764"), (5, 0.7999816888943144, "#FCE4D6"),
                 (5, -0.249977111117893, "#C65911"), (0, -0.1499984740745262, "#D9D9D9"),
                 (0, -0.0499893185216834, "#F2F2F2"), (0, -0.3499862666707358, "#A6A6A6"),
                 (1, 0.499984740745262, "#808080"), (1, 0.1499984740745262, "#262626"), (5, 0, "#ED7D31")]
        for theme, tint, expected in cases:
            with self.subTest(theme=theme, tint=tint):
                got = self.palette.cell(color(theme=theme, tint=tint))
                self.assertRegex(got, "^#[0-9A-F]{6}$")
                self.assertLessEqual(max(abs(int(got[i:i + 2], 16) - int(expected[i:i + 2], 16))
                                         for i in (1, 3, 5)), 2)

    def test_argb_indexed_and_automatic(self):
        self.assertEqual(self.palette.cell(color(rgb="FF445566")), "#445566")
        self.assertEqual(self.palette.cell(color(indexed=10)), "#FF0000")
        self.assertIsNone(self.palette.cell(color(indexed=64)))
        self.assertIsNone(self.palette.cell(color(auto=1)))
        self.assertIsNone(self.palette.cell(color(theme="x")))
        self.assertIsNone(self.palette.cell(None))

    def test_workbook_theme_and_custom_palette(self):
        from pp.template.colors import Palette, indexed_palette, theme_colors

        buffer = io.BytesIO()
        with zipfile.ZipFile(buffer, "w") as archive:
            archive.writestr("xl/_rels/workbook.xml.rels", (
                '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
                '<Relationship Id="rId9" Type="http://schemas.openxmlformats.org/officeDocument/2006/'
                'relationships/theme" Target="theme/theme1.xml"/></Relationships>'))
            archive.writestr("xl/theme/theme1.xml", (
                f'<a:theme xmlns:a="{A}"><a:themeElements><a:clrScheme name="t">'
                '<a:dk1><a:sysClr val="windowText" lastClr="111111"/></a:dk1>'
                '<a:lt1><a:sysClr val="window" lastClr="EEEEEE"/></a:lt1>'
                '<a:accent1><a:srgbClr val="123456"/></a:accent1>'
                '</a:clrScheme></a:themeElements></a:theme>'))
        styles = ET.fromstring(f'<styleSheet xmlns="{NS}"><colors><indexedColors>'
                               '<rgbColor rgb="FFABCDEF"/></indexedColors></colors></styleSheet>')
        with zipfile.ZipFile(buffer) as archive:
            palette = Palette(theme_colors(archive), indexed_palette(styles))
        self.assertEqual(palette.cell(color(theme=0)), "#EEEEEE")
        self.assertEqual(palette.cell(color(theme=1)), "#111111")
        self.assertEqual(palette.cell(color(theme=4)), "#123456")
        self.assertEqual(palette.cell(color(theme=5)), "#ED7D31")  # missing slot keeps the Office default
        self.assertEqual(palette.cell(color(indexed=0)), "#ABCDEF")


class DrawingColorTests(unittest.TestCase):
    def test_scheme_colors_with_luminance_changes(self):
        from pp.template.colors import Palette

        def fill(inner):
            return ET.fromstring(f'<a:solidFill xmlns:a="{A}">{inner}</a:solidFill>')

        palette = Palette()
        self.assertEqual(palette.drawing(fill('<a:srgbClr val="00B050"/>')), "#00B050")
        self.assertEqual(palette.drawing(fill('<a:schemeClr val="accent1"/>')), "#4472C4")
        # "Blue, Accent 1, Lighter 80%" as DrawingML writes it
        lighter = '<a:schemeClr val="accent1"><a:lumMod val="20000"/><a:lumOff val="80000"/></a:schemeClr>'
        self.assertEqual(palette.drawing(fill(lighter)), "#DAE3F3")
        self.assertEqual(palette.drawing(fill('<a:schemeClr val="tx1"/>')), "#000000")
        self.assertIsNone(palette.drawing(fill('<a:prstClr val="black"/>')))
        self.assertIsNone(palette.drawing(None))


if __name__ == "__main__":
    unittest.main()
