"""The static dependency index never misses a formula an edit can reach."""

import tempfile
import unittest
from pathlib import Path

from openpyxl import Workbook
from openpyxl.workbook.defined_name import DefinedName

from tests.v2.support import TMP


def build(formulas, sheets=("Calc", "Data Src"), names=()):
    from pp.engine.depindex import DependencyIndex

    TMP.mkdir(parents=True, exist_ok=True)
    directory = tempfile.TemporaryDirectory(dir=TMP)
    path = Path(directory.name) / "names.xlsx"
    book = Workbook()
    book.active.title = sheets[0]
    for name in sheets[1:]:
        book.create_sheet(name)
    for name, target in names:
        book.defined_names[name] = DefinedName(name, attr_text=target)
    book.save(path)
    index = DependencyIndex.build(path, formulas, list(sheets))
    directory.cleanup()
    return index


class ConeTests(unittest.TestCase):
    def test_chain_ranges_and_whole_columns(self):
        index = build({(0, 0, 1): "=A1*2", (0, 0, 2): "=B1+1", (0, 0, 3): "=SUM(B:B)",
                       (0, 0, 4): "=5", (0, 1, 1): "=SUM(A1:A3)"})
        self.assertEqual(index.cone([(0, 0, 0)]), {(0, 0, 1), (0, 0, 2), (0, 0, 3), (0, 1, 1)})
        self.assertEqual(index.cone([(0, 2, 0)]), {(0, 1, 1), (0, 0, 3)})
        self.assertEqual(index.cone([(0, 5, 0)]), set())

    def test_cross_sheet_quoted_names_and_strings(self):
        index = build({(0, 0, 0): "='Data Src'!B2*2", (0, 0, 1): '="Data Src!B2"&A1',
                       (1, 3, 0): "=SUM(Calc!A1:A1)"})
        self.assertEqual(index.cone([(1, 1, 1)]), {(0, 0, 0), (0, 0, 1), (1, 3, 0)})

    def test_defined_names_resolve_to_their_ranges(self):
        index = build({(0, 0, 0): "=Rate*2", (0, 0, 1): "=SUM(Block)"},
                      names=[("Rate", "'Data Src'!$B$2"), ("Block", "'Data Src'!$A$1:$C$9")])
        self.assertEqual(index.cone([(1, 1, 1)]), {(0, 0, 0), (0, 0, 1)})
        self.assertEqual(index.cone([(1, 5, 0)]), {(0, 0, 1)})

    def test_inputs_beyond_the_formula_area_are_seen(self):
        index = build({(0, 0, 0): "=SUM(Z:Z)", (0, 0, 1): "=Z5000", (0, 0, 2): "=SUM(4:4)"})
        self.assertEqual(index.cone([(0, 4999, 25)]), {(0, 0, 0), (0, 0, 1)})
        self.assertEqual(index.cone([(0, 3, 200)]), {(0, 0, 2)})

    def test_unknowable_and_volatile_formulas_are_always_affected(self):
        index = build({(0, 0, 0): "=OFFSET(A5,1,1)", (0, 0, 1): "=TODAY()", (0, 0, 2): "=C9"})
        self.assertEqual(index.cone([(0, 40, 40)]), {(0, 0, 0), (0, 0, 1)})


if __name__ == "__main__":
    unittest.main()
