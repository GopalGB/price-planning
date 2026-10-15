"""The compiled formula translator writes exactly what openpyxl's Translator writes, errors included."""

import random
import unittest

from openpyxl.formula.translate import Translator, TranslatorError
from openpyxl.utils.cell import get_column_letter

from pp.fidelity.conditional_formatting import _FastTranslator

FORMULAS = ("=A1+$B$2*C$3+$D4", "=SUM(A1:B5)+'My Sheet'!C3:D$9", '=AND($A2<>"",B:B>0,3:$4=1)',
            "=IF(MyName>0,Sheet2!$A$1,A1:MyName)", '="text with A1 inside"&B2', "=COUNTIF($A$1:A1,A1)>1",
            "plain", "", "=Z100*AA1")


def _written(function, *args):
    try:
        return function(*args)
    except TranslatorError:
        return "out of range"


class FastTranslatorTests(unittest.TestCase):
    def test_matches_openpyxl_for_every_target(self):
        rng = random.Random(1)
        for formula in FORMULAS:
            for origin in ("A1", "C5", "AA10"):
                slow = Translator(formula, origin=origin)
                fast = _FastTranslator(slow)
                for _ in range(200):
                    row, col = rng.randrange(0, 5000), rng.randrange(0, 60)
                    target = f"{get_column_letter(col + 1)}{row + 1}"
                    self.assertEqual(_written(fast.translate, row, col), _written(slow.translate_formula, target),
                                     (formula, origin, target))


if __name__ == "__main__":
    unittest.main()
