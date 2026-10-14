"""Tests for pp/export/extras_graft.py: extra sheets and extra columns grafted onto a template package.

Fixtures are built with openpyxl; a few parts are rewritten raw to plant shared formulas and calcPr.
"""
import re
import tempfile
import unittest
import zipfile
from pathlib import Path

from openpyxl import Workbook, load_workbook
from openpyxl.formatting.rule import CellIsRule
from openpyxl.styles import Font, PatternFill
from openpyxl.workbook.defined_name import DefinedName

from pp.export.extras_graft import GraftPlan, graft_extras, reorder_sheets

PLAN_SHEET = 'xl/worksheets/sheet1.xml'
GRAFTED = {'xl/workbook.xml', 'xl/_rels/workbook.xml.rels', '[Content_Types].xml',
           'xl/styles.xml', PLAN_SHEET}


def _rewrite(path, member, fn):
    """Fixture surgery: replace one zip member's text in place."""
    tmp = Path(str(path) + '.tmp')
    with zipfile.ZipFile(path) as zin, zipfile.ZipFile(tmp, 'w', zipfile.ZIP_DEFLATED) as zout:
        for info in zin.infolist():
            data = zin.read(info)
            if info.filename == member:
                data = fn(data.decode('utf-8')).encode('utf-8')
            zout.writestr(info.filename, data)
    tmp.replace(path)


def _set_cell(xml, ref, cell_xml):
    new, count = re.subn(rf'<c r="{ref}"(?: [^>]*?)?(?:/>|>.*?</c>)', cell_xml, xml, count=1, flags=re.S)
    assert count == 1, ref
    return new


def _template(path, cols='A'):
    """Plan (values, shared formula B3:B4, plain formula D3, styles) and Lookup with a sheet-scoped name."""
    wb = Workbook()
    plan = wb.active
    plan.title = 'Plan'
    rows = [['Item', 'Qty', 'Price', 'Note'], ['apple', 10, 2.5, 'x'], ['pear', 0, 1234.5, '=B3+1'], ['plum', 0]]
    for r, values in enumerate(rows, start=1):
        for c, value in enumerate(values, start=1):
            if value is not None:
                plan.cell(row=r, column=c, value=value)
    plan['C2'].font = Font(bold=True)
    plan['C2'].fill = PatternFill('solid', fgColor='FFFF0000')
    plan['C3'].number_format = '#,##0.00'
    if cols == 'A':
        plan.column_dimensions['A'].width = 22
    elif cols == 'group':
        plan.column_dimensions['A'].min, plan.column_dimensions['A'].max = 1, 10
        plan.column_dimensions['A'].width = 12
    lookup = wb.create_sheet('Lookup')
    lookup['A1'] = 'Key'
    lookup['A2'] = 'k1'
    wb.defined_names['TaxRate'] = DefinedName('TaxRate', attr_text='Plan!$C$2')
    lookup.defined_names['LookupLocal'] = DefinedName('LookupLocal', attr_text='Lookup!$A$2')
    wb.save(path)

    def plant(xml):
        xml = _set_cell(xml, 'B3', '<c r="B3"><f t="shared" ref="B3:B4" si="0">B2*2</f><v>20</v></c>')
        return _set_cell(xml, 'B4', '<c r="B4"><f t="shared" si="0"/><v>0</v></c>')

    _rewrite(path, PLAN_SHEET, plant)
    _rewrite(path, 'xl/workbook.xml', lambda xml: re.sub(r'<calcPr\b[^>]*/>', '<calcPr calcId="191029"/>', xml))


def _upload(path, e_width=15):
    """Same rows 1-4 as the template except B2=99, plus extra columns E:F, rows 5-7, and a My calc sheet."""
    wb = Workbook()
    plan = wb.active
    plan.title = 'Plan'
    rows = [['Item', 'Qty', 'Price', 'Note'], ['apple', 99, 2.5, 'x'], ['pear', 0, 1234.5, None], ['plum', 0]]
    for r, values in enumerate(rows, start=1):
        for c, value in enumerate(values, start=1):
            if value is not None:
                plan.cell(row=r, column=c, value=value)
    plan['E1'] = 'Score'
    plan['F1'] = 'Doubled'
    plan['E2'] = 0.25
    plan['E2'].number_format = '0.0%'
    plan['E2'].fill = PatternFill('solid', fgColor='FFFFFF00')
    plan['E3'] = 0.5
    plan['E4'] = 'merge'
    plan['F4'] = '=B4*2'
    plan['E6'] = 'beyond'
    plan['E7'] = 1.5
    for ref in ('F2', 'F3', 'B5', 'C5', 'D5', 'E5', 'F5'):
        plan[ref] = 0
    plan.column_dimensions['E'].width = e_width

    calc = wb.create_sheet('My calc')
    calc['A1'] = 'Total'
    calc['A1'].font = Font(bold=True, color='FF0000FF')
    calc['B1'] = '=Plan!B2+1'
    calc['C1'] = 1234.5
    calc['C1'].number_format = '0.00'
    calc.column_dimensions['C'].width = 30
    calc['A2'] = 'merged'
    calc.merge_cells('A2:B2')
    calc.conditional_formatting.add(
        'C1:C5', CellIsRule(operator='greaterThan', formula=['10'], font=Font(color='FFFF0000')))
    wb.defined_names['Bonus'] = DefinedName('Bonus', attr_text="'My calc'!$C$1")
    wb.defined_names['TaxRate'] = DefinedName('TaxRate', attr_text='Plan!$C$9')
    wb.defined_names['Broken'] = DefinedName('Broken', attr_text='#REF!')
    calc.defined_names['LocalX'] = DefinedName('LocalX', attr_text="'My calc'!$A$1")
    wb.save(path)

    def plant(xml):
        # Shared formula block si=0 with master in E2 (copied), dependent F3 (copied): renumbered.
        xml = _set_cell(xml, 'F2', '<c r="F2"><f t="shared" ref="F2:F3" si="0">E2*4</f><v>1</v></c>')
        xml = _set_cell(xml, 'F3', '<c r="F3"><f t="shared" si="0"/><v>2</v></c>')
        # Block si=1 with master in B5 (not copied): dependents E5/F5 lose their formula, keep cache.
        xml = _set_cell(xml, 'B5', '<c r="B5"><f t="shared" ref="B5:F5" si="1">B4*2</f><v>0</v></c>')
        xml = _set_cell(xml, 'C5', '<c r="C5"><f t="shared" si="1"/><v>1</v></c>')
        xml = _set_cell(xml, 'D5', '<c r="D5"><f t="shared" si="1"/><v>2</v></c>')
        xml = _set_cell(xml, 'E5', '<c r="E5"><f t="shared" si="1"/><v>3</v></c>')
        return _set_cell(xml, 'F5', '<c r="F5"><f t="shared" si="1"/><v>4</v></c>')

    _rewrite(path, PLAN_SHEET, plant)


def _sheet_xml(path, part):
    with zipfile.ZipFile(path) as z:
        return z.read(part).decode('utf-8')


class GraftTests(unittest.TestCase):
    def setUp(self):
        self._tmp = tempfile.TemporaryDirectory()
        self.dir = Path(self._tmp.name)
        self.template = self.dir / 'template.xlsx'
        self.upload = self.dir / 'upload.xlsx'
        self.out = self.dir / 'out.xlsx'
        _template(self.template)
        _upload(self.upload)

    def tearDown(self):
        self._tmp.cleanup()

    def graft(self, **kwargs):
        plan = GraftPlan(sheets=('My calc',), columns={'Plan': 4})
        return graft_extras(self.template, self.upload, kwargs.get('plan', plan), self.out)

    def test_graft_adds_sheet_and_columns_with_styles_formulas_and_names(self):
        counts = self.graft()
        self.assertEqual(counts['sheets'], 1)
        self.assertEqual(counts['dropped_formulas'], 0)
        wb = load_workbook(self.out)
        self.assertEqual(wb.sheetnames, ['Plan', 'Lookup', 'My calc'])

        template = load_workbook(self.template)
        for row in template['Plan']['A1:D4']:
            for cell in row:
                self.assertEqual(wb['Plan'][cell.coordinate].value, cell.value, cell.coordinate)
        plan = wb['Plan']
        self.assertEqual(plan['B2'].value, 10)
        self.assertIsNone(plan['B5'].value)
        self.assertTrue(plan['C2'].font.b)
        self.assertEqual(plan['C2'].fill.fgColor.rgb, 'FFFF0000')
        self.assertEqual(plan['C3'].number_format, '#,##0.00')
        self.assertEqual(plan.column_dimensions['A'].width, 22)

        self.assertEqual(plan['E1'].value, 'Score')
        self.assertEqual(plan['E2'].value, 0.25)
        self.assertEqual(plan['E2'].number_format, '0.0%')
        self.assertEqual(plan['E2'].fill.fgColor.rgb, 'FFFFFF00')
        self.assertEqual(plan['E4'].value, 'merge')
        self.assertEqual(plan['F2'].value, '=E2*4')
        self.assertEqual(plan['F3'].value, '=E3*4')
        self.assertEqual(plan['F4'].value, '=B4*2')
        self.assertEqual(plan['E6'].value, 'beyond')
        self.assertEqual(plan['E7'].value, 1.5)
        self.assertEqual(plan.column_dimensions['E'].width, 15)
        self.assertEqual(plan['E5'].value, '=E4*2')  # shared children whose master stays left get their own formula
        self.assertEqual(plan['F5'].value, '=F4*2')
        self.assertIn('<dimension ref="A1:F7"', _sheet_xml(self.out, PLAN_SHEET))

        calc = wb['My calc']
        self.assertEqual(calc['A1'].value, 'Total')
        self.assertTrue(calc['A1'].font.b)
        self.assertEqual(calc['A1'].font.color.rgb, 'FF0000FF')
        self.assertEqual(calc['B1'].value, '=Plan!B2+1')
        self.assertEqual(calc['C1'].value, 1234.5)
        self.assertEqual(calc['C1'].number_format, '0.00')
        self.assertEqual(calc.column_dimensions['C'].width, 30)
        self.assertIn('A2:B2', [str(r) for r in calc.merged_cells.ranges])
        rules = [rule for cf in calc.conditional_formatting for rule in cf.rules]
        self.assertEqual(rules[0].dxf.font.color.rgb, 'FFFF0000')

        self.assertIn('Bonus', wb.defined_names)
        self.assertEqual(wb.defined_names['TaxRate'].attr_text, 'Plan!$C$2')
        self.assertNotIn('Broken', wb.defined_names)
        self.assertNotIn('LocalX', wb.defined_names)

        with zipfile.ZipFile(self.out) as z:
            self.assertNotIn('xl/calcChain.xml', z.namelist())
            self.assertIn('fullCalcOnLoad="1"', z.read('xl/workbook.xml').decode('utf-8'))

    def test_shared_children_of_a_master_left_behind_get_their_own_formula(self):
        self.graft()
        xml = _sheet_xml(self.out, PLAN_SHEET)
        for ref, formula in (('E5', 'E4*2'), ('F5', 'F4*2')):
            cell = re.search(rf'<c r="{ref}"[^>]*>.*?</c>', xml, re.S).group(0)
            self.assertIn(f'<f>{formula}</f>', cell)
            self.assertIn('<v>', cell)

    def test_untouched_parts_are_copied_byte_for_byte(self):
        self.graft()
        with zipfile.ZipFile(self.template) as t, zipfile.ZipFile(self.out) as o:
            for name in t.namelist():
                if name not in GRAFTED:
                    self.assertEqual(o.read(name), t.read(name), name)

    def test_refuses_to_write_over_template(self):
        with self.assertRaises(ValueError):
            graft_extras(self.template, self.upload, GraftPlan(sheets=(), columns={}), self.template)

    def test_column_widths_split_template_ranges(self):
        _template(self.template, cols='group')
        self.graft()
        xml = _sheet_xml(self.out, PLAN_SHEET)
        cols = []
        for tag in re.findall(r'<col\b[^>]*/>', xml):
            lo, hi = int(re.search(r'\smin="(\d+)"', tag).group(1)), int(re.search(r'\smax="(\d+)"', tag).group(1))
            cols.append((lo, hi, float(re.search(r'\swidth="([\d.]+)"', tag).group(1))))
        self.assertEqual(cols, [(1, 4, 12.0), (5, 5, 15.0), (6, 10, 12.0)])

    def test_cols_container_created_after_sheet_format_pr(self):
        _template(self.template, cols=None)
        self.graft()
        xml = _sheet_xml(self.out, PLAN_SHEET)
        self.assertLess(xml.index('<sheetFormatPr'), xml.index('<cols>'))
        self.assertLess(xml.index('<cols>'), xml.index('<sheetData'))
        self.assertIn('width="15"', xml)

    def test_reorder_moves_sheet_and_remaps_local_names(self):
        self.graft()
        reorder_sheets(self.out, ('Plan', 'My calc'))
        wb = load_workbook(self.out)
        self.assertEqual(wb.sheetnames, ['Plan', 'My calc', 'Lookup'])
        self.assertIn('LookupLocal', wb['Lookup'].defined_names)
        self.assertNotIn('LookupLocal', wb['My calc'].defined_names)
        reorder_sheets(self.out, ('Lookup',))
        wb = load_workbook(self.out)
        self.assertEqual(wb.sheetnames, ['Lookup', 'Plan', 'My calc'])
        self.assertIn('LookupLocal', wb['Lookup'].defined_names)

    def test_reorder_matches_names_case_insensitively(self):
        self.graft()
        reorder_sheets(self.out, ('my CALC', 'lookup'))
        self.assertEqual(load_workbook(self.out).sheetnames, ['My calc', 'Lookup', 'Plan'])
        self.assertIn('LookupLocal', load_workbook(self.out)['Lookup'].defined_names)

    def test_reorder_ignores_names_missing_from_package(self):
        self.graft()
        reorder_sheets(self.out, ('Gone', 'My calc', 'Also gone'))
        wb = load_workbook(self.out)
        self.assertEqual(wb.sheetnames, ['My calc', 'Plan', 'Lookup'])
        self.assertIn('LookupLocal', wb['Lookup'].defined_names)


if __name__ == '__main__':
    unittest.main()
