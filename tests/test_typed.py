# tests/test_typed.py
import sys, unittest
from pathlib import Path
sys.path.insert(0, str(Path(__file__).resolve().parent.parent / 'ops'))
from typed import typed, same

class Typed(unittest.TestCase):
    def test_typed_compare_distinguishes_blank_empty_zero_error(self):
        kinds = [typed(None)[0], typed('')[0], typed(0)[0], typed('#N/A')[0], typed(True)[0], typed('abc')[0]]
        self.assertEqual(kinds, ['blank', 'empty', 'num', 'error', 'bool', 'text'])
        self.assertFalse(same(None, ''))
        self.assertFalse(same('', 0))
        self.assertFalse(same(0, '#N/A'))
    def test_numeric_tolerance(self):
        self.assertTrue(same(1.0, 1.0 + 1e-12))
        self.assertFalse(same(1.0, 1.001))
    def test_bool_is_not_number(self):
        self.assertFalse(same(True, 1))
