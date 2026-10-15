"""Tests for the Excel value model in pp.calc.values."""

import unittest

from pp.calc.values import (
    CALC,
    DIV0,
    NA,
    NAME,
    NULL,
    NUM,
    REF,
    SPILL,
    VALUE,
    Area,
    XlError,
    compare,
    first_error,
    scalar_of,
    to_bool,
    to_number,
    to_text,
)


class XlErrorTests(unittest.TestCase):
    def test_equal_by_code(self):
        self.assertEqual(XlError("#DIV/0!"), DIV0)
        self.assertNotEqual(XlError("#N/A"), VALUE)
        self.assertNotEqual(NA, "#N/A")

    def test_hashable(self):
        self.assertEqual(len({XlError("#N/A"), NA, VALUE}), 2)

    def test_repr_mentions_code(self):
        self.assertIn("#N/A", repr(NA))
        self.assertEqual(NA.code, "#N/A")

    def test_constants_carry_excel_codes(self):
        expected = {
            NA: "#N/A", VALUE: "#VALUE!", REF: "#REF!", DIV0: "#DIV/0!",
            NUM: "#NUM!", NAME: "#NAME?", NULL: "#NULL!", SPILL: "#SPILL!",
            CALC: "#CALC!",
        }
        for err, code in expected.items():
            self.assertEqual(err.code, code)

    def test_unknown_code_rejected(self):
        with self.assertRaises(ValueError):
            XlError("#FOO!")


class AreaTests(unittest.TestCase):
    def test_shape_and_get(self):
        a = Area([[1, 2, 3], [4, 5, 6]])
        self.assertEqual((a.height, a.width), (2, 3))
        self.assertEqual(a.get(1, 2), 6)

    def test_ints_normalised_to_float_but_bools_kept(self):
        a = Area([[1, True, None, "x"]])
        self.assertIs(type(a.get(0, 0)), float)
        self.assertIs(a.get(0, 1), True)
        self.assertIsNone(a.get(0, 2))
        self.assertEqual(a.get(0, 3), "x")

    def test_cells_row_major(self):
        self.assertEqual(list(Area([[1, 2], [3, 4]]).cells()), [1.0, 2.0, 3.0, 4.0])

    def test_column_and_row(self):
        a = Area([[1, 2], [3, 4]])
        self.assertEqual(a.column(1), [2.0, 4.0])
        self.assertEqual(a.row(0), [1.0, 2.0])

    def test_ragged_and_empty_rejected(self):
        with self.assertRaises(ValueError):
            Area([[1, 2], [3]])
        with self.assertRaises(ValueError):
            Area([])
        with self.assertRaises(ValueError):
            Area([[]])

    def test_nested_area_and_objects_rejected(self):
        with self.assertRaises(TypeError):
            Area([[Area.scalar(1)]])
        with self.assertRaises(TypeError):
            Area([[object()]])

    def test_immutable_source_and_no_item_assignment(self):
        rows = [[1]]
        a = Area(rows)
        rows[0][0] = 9
        self.assertEqual(a.get(0, 0), 1.0)
        self.assertFalse(hasattr(a, "__setitem__"))
        col = a.column(0)
        col.append(7)
        self.assertEqual(a.height, 1)

    def test_memo_is_persistent_per_area(self):
        a = Area([[1]])
        self.assertEqual(a.memo, {})
        a.memo["k"] = 1
        self.assertIs(a.memo, a.memo)
        self.assertEqual(a.memo["k"], 1)
        self.assertEqual(Area([[1]]).memo, {})

    def test_equality_by_content(self):
        self.assertEqual(Area([[1, 2]]), Area([[1.0, 2.0]]))
        self.assertNotEqual(Area([[1, 2]]), Area([[2, 1]]))

    def test_scalar_helper(self):
        a = Area.scalar("x")
        self.assertEqual((a.height, a.width), (1, 1))
        self.assertEqual(a.get(0, 0), "x")

    def test_scalar_of(self):
        self.assertEqual(scalar_of(3), 3)
        self.assertIsNone(scalar_of(None))
        self.assertEqual(scalar_of(Area.scalar(2)), 2.0)
        self.assertEqual(scalar_of(Area([[1, 2]])), VALUE)


class ToNumberTests(unittest.TestCase):
    def test_blank_and_bool(self):
        self.assertEqual(to_number(None), 0.0)
        self.assertEqual(to_number(True), 1.0)
        self.assertEqual(to_number(False), 0.0)

    def test_numbers_pass_through(self):
        self.assertEqual(to_number(3), 3.0)
        self.assertEqual(to_number(-2.5), -2.5)

    def test_numeric_text(self):
        self.assertEqual(to_number("  12.5 "), 12.5)
        self.assertEqual(to_number("1,234.5"), 1234.5)
        self.assertEqual(to_number("50%"), 0.5)
        self.assertEqual(to_number("1e3"), 1000.0)
        self.assertEqual(to_number("-5"), -5.0)
        self.assertEqual(to_number("+7"), 7.0)
        self.assertEqual(to_number(".5"), 0.5)
        self.assertEqual(to_number("5."), 5.0)

    def test_non_numeric_text_is_value_error(self):
        self.assertEqual(to_number("abc"), VALUE)
        self.assertEqual(to_number(""), VALUE)
        self.assertEqual(to_number("   "), VALUE)
        self.assertEqual(to_number("TRUE"), VALUE)
        self.assertEqual(to_number("1.2.3"), VALUE)

    def test_errors_pass_through(self):
        self.assertEqual(to_number(DIV0), DIV0)

    def test_area_coerced_through_scalar_of(self):
        self.assertEqual(to_number(Area.scalar("4")), 4.0)

    def test_overflow_text_is_not_a_finite_number(self):
        self.assertEqual(to_number("1e999"), VALUE)


class ToTextTests(unittest.TestCase):
    def test_blank_bool_text(self):
        self.assertEqual(to_text(None), "")
        self.assertEqual(to_text(True), "TRUE")
        self.assertEqual(to_text(False), "FALSE")
        self.assertEqual(to_text("hi"), "hi")

    def test_integers_have_no_decimal_point(self):
        self.assertEqual(to_text(5.0), "5")
        self.assertEqual(to_text(100.0), "100")
        self.assertEqual(to_text(0.0), "0")
        self.assertEqual(to_text(-0.0), "0")
        self.assertEqual(to_text(123456789012.0), "123456789012")

    def test_fifteen_significant_digits(self):
        self.assertEqual(to_text(0.1 + 0.2), "0.3")
        self.assertEqual(to_text(1 / 3), "0.333333333333333")
        self.assertEqual(to_text(2 / 3), "0.666666666666667")
        self.assertEqual(to_text(1234.5678), "1234.5678")
        self.assertEqual(to_text(-3.5), "-3.5")
        self.assertEqual(to_text(0.1), "0.1")

    def test_scientific_for_huge_and_tiny(self):
        self.assertEqual(to_text(1e21), "1E+21")
        self.assertEqual(to_text(1234567890123456789.0), "1.23456789012346E+18")
        self.assertEqual(to_text(1e-7), "1E-07")

    def test_small_fixed_notation_kept(self):
        self.assertEqual(to_text(0.000001), "0.000001")
        self.assertEqual(to_text(1.5e-5), "0.000015")

    def test_area_coerced(self):
        self.assertEqual(to_text(Area.scalar(7.0)), "7")


class ToBoolTests(unittest.TestCase):
    def test_numbers_and_blank(self):
        self.assertIs(to_bool(None), False)
        self.assertIs(to_bool(0.0), False)
        self.assertIs(to_bool(2.0), True)
        self.assertIs(to_bool(True), True)

    def test_text(self):
        self.assertIs(to_bool("true"), True)
        self.assertIs(to_bool("FALSE"), False)
        self.assertEqual(to_bool("yes"), VALUE)

    def test_error_and_area(self):
        self.assertEqual(to_bool(NA), NA)
        self.assertIs(to_bool(Area.scalar(1)), True)


class FirstErrorTests(unittest.TestCase):
    def test_none_when_clean(self):
        self.assertIsNone(first_error(1.0, "a", None))
        self.assertIsNone(first_error())

    def test_first_in_argument_order(self):
        self.assertEqual(first_error(1.0, NA, DIV0), NA)

    def test_scans_areas_row_major(self):
        self.assertEqual(first_error(Area([[1, DIV0], [NA, 2]])), DIV0)
        self.assertEqual(first_error(Area([[1], [NA]]), DIV0), NA)


class CompareTests(unittest.TestCase):
    def test_numbers(self):
        self.assertEqual(compare(1, 2), -1)
        self.assertEqual(compare(2.0, 2.0), 0)
        self.assertEqual(compare(3.0, 1.0), 1)

    def test_type_order_number_text_bool(self):
        self.assertEqual(compare(3.0, "a"), -1)
        self.assertEqual(compare("a", True), -1)
        self.assertEqual(compare(True, 5.0), 1)
        self.assertEqual(compare(1.0, True), -1)

    def test_text_case_insensitive(self):
        self.assertEqual(compare("ABC", "abc"), 0)
        self.assertEqual(compare("apple", "Banana"), -1)

    def test_blank_takes_the_other_type_zero_value(self):
        self.assertEqual(compare(None, 0.0), 0)
        self.assertEqual(compare(None, -1.0), 1)
        self.assertEqual(compare(None, ""), 0)
        self.assertEqual(compare("", None), 0)
        self.assertEqual(compare(None, False), 0)
        self.assertEqual(compare(None, None), 0)
        self.assertEqual(compare(None, "a"), -1)


if __name__ == "__main__":
    unittest.main()
