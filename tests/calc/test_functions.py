"""Tests for the Excel function library in pp.calc.functions."""

import time
import unittest

from pp.calc.functions import FUNCTIONS, LAZY
from pp.calc.values import (
    CALC,
    DIV0,
    NA,
    NUM,
    REF,
    VALUE,
    Area,
)


def call(name, *args):
    return FUNCTIONS[name](*args)


TABLE = Area([[1, "one", 10], [2, "two", 20], [4, "four", 40]])


class TextFunctionTests(unittest.TestCase):
    def test_substitute(self):
        self.assertEqual(call("SUBSTITUTE", "Hello World", "o", "0"), "Hell0 W0rld")
        self.assertEqual(call("SUBSTITUTE", "aaa", "a", "b", 2), "aba")
        self.assertEqual(call("SUBSTITUTE", "aaa", "a", "b", 9), "aaa")
        self.assertEqual(call("SUBSTITUTE", "aaa", "a", "b", 0), VALUE)
        self.assertEqual(call("SUBSTITUTE", "aaa", "aa", "b", 2), "aaa")
        self.assertEqual(call("SUBSTITUTE", "abc", "", "x"), "abc")
        self.assertEqual(call("SUBSTITUTE", "Abc", "a", "x"), "Abc")
        self.assertEqual(call("SUBSTITUTE", "a", "a", "b", 1.9), "b")

    def test_substitute_error_propagation(self):
        self.assertEqual(call("SUBSTITUTE", "x", DIV0, "y"), DIV0)
        self.assertEqual(call("SUBSTITUTE", "x", "x", "y", NA), NA)

    def test_trim_collapses_inner_spaces_only(self):
        self.assertEqual(call("TRIM", "  a   b  c "), "a b c")
        self.assertEqual(call("TRIM", "a\tb"), "a\tb")
        self.assertEqual(call("TRIM", None), "")
        self.assertEqual(call("TRIM", NA), NA)
        self.assertEqual(call("TRIM", 12.0), "12")

    def test_numbervalue(self):
        self.assertEqual(call("NUMBERVALUE", "1.234,5", ",", "."), 1234.5)
        self.assertEqual(call("NUMBERVALUE", "1,234.5"), 1234.5)
        self.assertEqual(call("NUMBERVALUE", "50%"), 0.5)
        self.assertEqual(call("NUMBERVALUE", ""), 0.0)
        self.assertEqual(call("NUMBERVALUE", "1 000"), 1000.0)
        self.assertEqual(call("NUMBERVALUE", "abc"), VALUE)
        self.assertEqual(call("NUMBERVALUE", "12", ",", ","), VALUE)

    def test_concat_flattens_areas(self):
        self.assertEqual(call("CONCAT", "a", 1.0, True), "a1TRUE")
        self.assertEqual(call("CONCAT", Area([["x", "y"]]), "z"), "xyz")
        self.assertEqual(call("CONCAT", None, "a"), "a")
        self.assertEqual(call("CONCAT", "a", NA), NA)
        self.assertEqual(call("CONCAT", Area([["p", DIV0]])), DIV0)

    def test_concatenate_lifts(self):
        self.assertEqual(call("CONCATENATE", "a", "b", 3.0), "ab3")
        self.assertEqual(call("CONCATENATE", Area([["a", "b"]])), Area([["a", "b"]]))

    def test_left(self):
        self.assertEqual(call("LEFT", "hello"), "h")
        self.assertEqual(call("LEFT", "hello", 3.9), "hel")
        self.assertEqual(call("LEFT", "hello", 0), "")
        self.assertEqual(call("LEFT", "hello", 99), "hello")
        self.assertEqual(call("LEFT", "hello", -1), VALUE)
        self.assertEqual(call("LEFT", 12345.0, 2), "12")
        self.assertEqual(call("LEFT", NA), NA)

    def test_right(self):
        self.assertEqual(call("RIGHT", "hello"), "o")
        self.assertEqual(call("RIGHT", "hello", 2), "lo")
        self.assertEqual(call("RIGHT", "hi", 5), "hi")
        self.assertEqual(call("RIGHT", "hi", 0), "")
        self.assertEqual(call("RIGHT", "x", -1), VALUE)

    def test_mid(self):
        self.assertEqual(call("MID", "hello", 2, 3), "ell")
        self.assertEqual(call("MID", "hello", 2, 99), "ello")
        self.assertEqual(call("MID", "hello", 9, 1), "")
        self.assertEqual(call("MID", "hello", 0, 1), VALUE)
        self.assertEqual(call("MID", "hello", 1, -1), VALUE)
        self.assertEqual(call("MID", DIV0, 1, 1), DIV0)

    def test_find_case_sensitive(self):
        self.assertEqual(call("FIND", "l", "hello"), 3.0)
        self.assertEqual(call("FIND", "L", "hello"), VALUE)
        self.assertEqual(call("FIND", "l", "hello", 4), 4.0)
        self.assertEqual(call("FIND", "", "abc"), 1.0)
        self.assertEqual(call("FIND", "l", "hello", 0), VALUE)
        self.assertEqual(call("FIND", "a", "abc", 9), VALUE)
        self.assertEqual(call("FIND", "x", "abc"), VALUE)
        self.assertEqual(call("FIND", "*", "a*b"), 2.0)  # FIND is literal, no wildcards

    def test_search_case_insensitive_with_wildcards(self):
        self.assertEqual(call("SEARCH", "L", "hello"), 3.0)
        self.assertEqual(call("SEARCH", "l*o", "hello"), 3.0)
        self.assertEqual(call("SEARCH", "h?llo", "HELLO"), 1.0)
        self.assertEqual(call("SEARCH", "~*", "a*b"), 2.0)
        self.assertEqual(call("SEARCH", "~?", "a?"), 2.0)
        self.assertEqual(call("SEARCH", "l", "hello", 4), 4.0)
        self.assertEqual(call("SEARCH", "x", "abc"), VALUE)

    def test_len(self):
        self.assertEqual(call("LEN", "hello"), 5.0)
        self.assertEqual(call("LEN", None), 0.0)
        self.assertEqual(call("LEN", True), 4.0)
        self.assertEqual(call("LEN", 12.5), 4.0)
        self.assertEqual(call("LEN", Area([["ab", "c"]])), Area([[2.0, 1.0]]))

    def test_value(self):
        self.assertEqual(call("VALUE", "1,234"), 1234.0)
        self.assertAlmostEqual(call("VALUE", "5%"), 0.05)
        self.assertEqual(call("VALUE", ""), VALUE)
        self.assertEqual(call("VALUE", None), 0.0)
        self.assertEqual(call("VALUE", True), VALUE)
        self.assertEqual(call("VALUE", "x"), VALUE)

    def test_text_number_formats(self):
        self.assertEqual(call("TEXT", 1234.5, "#,##0"), "1,235")
        self.assertEqual(call("TEXT", -1234.5, "#,##0"), "-1,235")
        self.assertEqual(call("TEXT", 1234.5678, "0.00"), "1234.57")
        self.assertEqual(call("TEXT", 0.256, "0%"), "26%")
        self.assertEqual(call("TEXT", 7, "00"), "07")
        self.assertEqual(call("TEXT", 7, "000"), "007")
        self.assertEqual(call("TEXT", 0.5, "0"), "1")
        self.assertEqual(call("TEXT", 0.1 + 0.2, "0.00"), "0.30")
        self.assertEqual(call("TEXT", "12", "0.00"), "12.00")
        self.assertEqual(call("TEXT", None, "0.00"), "0.00")

    def test_text_at_and_dates(self):
        self.assertEqual(call("TEXT", 12.0, "@"), "12")
        self.assertEqual(call("TEXT", 45000.0, "yyyy-mm-dd"), "2023-03-15")
        self.assertEqual(call("TEXT", 45000.0, "mm/dd/yyyy"), "03/15/2023")
        self.assertEqual(call("TEXT", 1.0, "yyyy-mm-dd"), "1900-01-01")
        self.assertEqual(call("TEXT", 61.0, "yyyy-mm-dd"), "1900-03-01")

    def test_text_fallbacks(self):
        self.assertEqual(call("TEXT", "abc", "0.00"), "abc")
        self.assertEqual(call("TEXT", DIV0, "0"), DIV0)
        self.assertEqual(call("TEXT", 5.0, "0.0 units"), "5")
        self.assertEqual(call("TEXT", True, "0"), "TRUE")


class MathFunctionTests(unittest.TestCase):
    def test_int_floors(self):
        self.assertEqual(call("INT", 5.7), 5.0)
        self.assertEqual(call("INT", -5.7), -6.0)
        self.assertEqual(call("INT", -1.5), -2.0)
        self.assertEqual(call("INT", "3.9"), 3.0)
        self.assertEqual(call("INT", NA), NA)
        self.assertEqual(call("INT", "x"), VALUE)
        self.assertEqual(call("INT", Area([[1.5, -1.5]])), Area([[1.0, -2.0]]))

    def test_floor_significance(self):
        self.assertEqual(call("FLOOR", 3.7, 2), 2.0)
        self.assertEqual(call("FLOOR", -2.5, 2), -4.0)
        self.assertEqual(call("FLOOR", -2.5, -2), -2.0)
        self.assertEqual(call("FLOOR", 2.5, 0), 0.0)
        self.assertEqual(call("FLOOR", 2.5, -2), NUM)
        self.assertEqual(call("FLOOR", 0.3, 0.1), 0.3)

    def test_floor_math(self):
        self.assertEqual(call("FLOOR.MATH", -2.5), -3.0)
        self.assertEqual(call("FLOOR.MATH", 6.7, 2), 6.0)
        self.assertEqual(call("FLOOR.MATH", -5.5, 2), -6.0)
        self.assertEqual(call("FLOOR.MATH", -5.5, 2, 1), -4.0)
        self.assertEqual(call("FLOOR.MATH", 5, 0), 0.0)
        self.assertEqual(call("FLOOR.MATH", 10), 10.0)

    def test_roundup_away_from_zero(self):
        self.assertEqual(call("ROUNDUP", 3.2001, 2), 3.21)
        self.assertEqual(call("ROUNDUP", -3.2001, 2), -3.21)
        self.assertEqual(call("ROUNDUP", 3.0, 0), 3.0)
        self.assertEqual(call("ROUNDUP", 1234.5, -2), 1300.0)
        self.assertEqual(call("ROUNDUP", NA, 0), NA)

    def test_mod_takes_divisor_sign(self):
        self.assertEqual(call("MOD", 10, 3), 1.0)
        self.assertEqual(call("MOD", -3, 2), 1.0)
        self.assertEqual(call("MOD", 3, -2), -1.0)
        self.assertEqual(call("MOD", 5, 0), DIV0)
        self.assertEqual(call("MOD", 5.5, 2), 1.5)
        self.assertEqual(call("MOD", -5.5, 2), 0.5)

    def test_sum_ignores_text_bool_blank_in_areas(self):
        self.assertEqual(call("SUM", 1.0, 2.0, 3.0), 6.0)
        self.assertEqual(call("SUM", Area([[1, "a", True, None, 2]])), 3.0)
        self.assertEqual(call("SUM", Area([[1, DIV0]])), DIV0)
        self.assertEqual(call("SUM", "5", 1), 6.0)
        self.assertEqual(call("SUM", "x"), VALUE)
        self.assertEqual(call("SUM", Area([[1, 2], [3, 4]]), 10), 20.0)
        self.assertEqual(call("SUM", True), 1.0)
        self.assertEqual(call("SUM", Area([[1, NA]])), NA)

    def test_max_min(self):
        self.assertEqual(call("MAX", Area([[1, "x", 5, None]])), 5.0)
        self.assertEqual(call("MAX", Area([["x"]])), 0.0)
        self.assertEqual(call("MAX", 3, "4"), 4.0)
        self.assertEqual(call("MAX", Area([[1, NA]])), NA)
        self.assertEqual(call("MIN", Area([[4, -2]])), -2.0)
        self.assertEqual(call("MIN", 3, Area([[1]])), 1.0)

    def test_sumproduct(self):
        self.assertEqual(call("SUMPRODUCT", Area([[1, 2]]), Area([[3, 4]])), 11.0)
        self.assertEqual(call("SUMPRODUCT", Area([[1], [2]]), Area([[3, 4]])), VALUE)
        self.assertEqual(call("SUMPRODUCT", Area([[1, "x"]]), Area([[2, 5]])), 2.0)
        self.assertEqual(call("SUMPRODUCT", Area([[1, NA]]), Area([[1, 1]])), NA)
        self.assertEqual(
            call("SUMPRODUCT", Area([[1, 2], [3, 4]]), Area([[1, 1], [1, 1]])), 10.0
        )

    def test_iseven(self):
        self.assertIs(call("ISEVEN", 4), True)
        self.assertIs(call("ISEVEN", 3), False)
        self.assertIs(call("ISEVEN", 2.9), True)
        self.assertIs(call("ISEVEN", -4), True)
        self.assertEqual(call("ISEVEN", "x"), VALUE)
        self.assertEqual(call("ISEVEN", Area([[1, 2]])), Area([[False, True]]))

    def test_days(self):
        self.assertEqual(call("DAYS", 10, 3), 7.0)
        self.assertEqual(call("DAYS", 45010.7, 45000.2), 10.0)
        self.assertEqual(call("DAYS", "45010", 45000), 10.0)
        self.assertEqual(call("DAYS", NA, 1), NA)


class InfoFunctionTests(unittest.TestCase):
    def test_iserror(self):
        self.assertIs(call("ISERROR", NA), True)
        self.assertIs(call("ISERROR", 1.0), False)
        self.assertIs(call("ISERROR", None), False)
        self.assertEqual(call("ISERROR", Area([[NA, 1]])), Area([[True, False]]))

    def test_isnumber(self):
        self.assertIs(call("ISNUMBER", 1.0), True)
        self.assertIs(call("ISNUMBER", "1"), False)
        self.assertIs(call("ISNUMBER", True), False)
        self.assertIs(call("ISNUMBER", NA), False)
        self.assertIs(call("ISNUMBER", None), False)

    def test_isblank(self):
        self.assertIs(call("ISBLANK", None), True)
        self.assertIs(call("ISBLANK", ""), False)
        self.assertIs(call("ISBLANK", 0.0), False)
        self.assertIs(call("ISBLANK", Area.scalar(None)), True)

    def test_istext(self):
        self.assertIs(call("ISTEXT", "a"), True)
        self.assertIs(call("ISTEXT", ""), True)
        self.assertIs(call("ISTEXT", 1.0), False)
        self.assertIs(call("ISTEXT", NA), False)

    def test_not(self):
        self.assertIs(call("NOT", True), False)
        self.assertIs(call("NOT", 0.0), True)
        self.assertIs(call("NOT", "TRUE"), False)
        self.assertEqual(call("NOT", "x"), VALUE)
        self.assertEqual(call("NOT", NA), NA)
        self.assertIs(call("NOT", None), True)

    def test_and(self):
        self.assertIs(call("AND", True, 1.0), True)
        self.assertIs(call("AND", True, 0), False)
        self.assertIs(call("AND", Area([[True, "x", None]])), True)
        self.assertEqual(call("AND", Area([["x", None]])), VALUE)
        self.assertEqual(call("AND", Area([[True, NA]])), NA)
        self.assertIs(call("AND", "TRUE"), True)
        self.assertEqual(call("AND", "x"), VALUE)

    def test_or(self):
        self.assertIs(call("OR", False, 0.0), False)
        self.assertIs(call("OR", Area([[False, 2.0]])), True)
        self.assertEqual(call("OR", Area([["x"]])), VALUE)
        self.assertEqual(call("OR", Area([[NA, True]])), NA)

    def test_counta(self):
        self.assertEqual(call("COUNTA", Area([[1, "", None, NA]])), 3.0)
        self.assertEqual(call("COUNTA", "a", None, 2.0), 2.0)


class LookupFunctionTests(unittest.TestCase):
    def test_match_approximate_ascending(self):
        self.assertEqual(call("MATCH", 3, Area([[1, 2, 3, 4]])), 3.0)
        self.assertEqual(call("MATCH", 2.5, Area([[1], [2], [3]]), 1), 2.0)
        self.assertEqual(call("MATCH", 0, Area([[1, 2, 3]]), 1), NA)
        self.assertEqual(call("MATCH", 10, Area([[1, 2, 3]])), 3.0)

    def test_match_approximate_descending(self):
        desc = Area([[9, 7, 5, 3]])
        self.assertEqual(call("MATCH", 5, desc, -1), 3.0)
        self.assertEqual(call("MATCH", 6, desc, -1), 2.0)
        self.assertEqual(call("MATCH", 4, desc, -1), 3.0)
        self.assertEqual(call("MATCH", 10, desc, -1), NA)

    def test_match_exact_case_insensitive_and_wildcards(self):
        self.assertEqual(call("MATCH", "b", Area([["a", "B", "c"]]), 0), 2.0)
        self.assertEqual(call("MATCH", "c*", Area([["apple", "cat", "cow"]]), 0), 2.0)
        self.assertEqual(call("MATCH", "?ow", Area([["bow", "cow", "wow"]]), 0), 1.0)
        self.assertEqual(call("MATCH", "x", Area([["a"]]), 0), NA)
        self.assertEqual(call("MATCH", "2", Area([[2.0]]), 0), NA)
        self.assertEqual(call("MATCH", DIV0, Area([[1]]), 0), DIV0)
        self.assertEqual(call("MATCH", 1, Area([[1, 2], [3, 4]]), 0), NA)

    def test_vlookup_exact(self):
        self.assertEqual(call("VLOOKUP", 2, TABLE, 2, False), "two")
        self.assertEqual(call("VLOOKUP", 3, TABLE, 2, False), NA)
        self.assertEqual(call("VLOOKUP", 4, TABLE, 3, False), 40.0)
        self.assertEqual(call("VLOOKUP", 2, TABLE, 4, False), REF)
        self.assertEqual(call("VLOOKUP", 2, TABLE, 0, False), VALUE)
        self.assertEqual(call("VLOOKUP", 2, TABLE, 1.9, False), 2.0)
        self.assertEqual(call("VLOOKUP", DIV0, TABLE, 2, False), DIV0)

    def test_vlookup_approximate_default(self):
        self.assertEqual(call("VLOOKUP", 3, TABLE, 2), "two")
        self.assertEqual(call("VLOOKUP", 5, TABLE, 2), "four")
        self.assertEqual(call("VLOOKUP", 0.5, TABLE, 2), NA)
        self.assertEqual(call("VLOOKUP", 3, TABLE, 2, True), "two")

    def test_vlookup_text_exact_wildcard_and_case(self):
        self.assertEqual(call("VLOOKUP", "TWO", Area([["two", "x"]]), 2, False), "x")
        self.assertEqual(
            call("VLOOKUP", "t*", Area([["two", 1], ["four", 2]]), 2, False), 1.0
        )

    def test_vlookup_blank_result_is_zero(self):
        self.assertEqual(call("VLOOKUP", 1, Area([[1, None]]), 2, False), 0.0)

    def test_hlookup(self):
        htable = Area([[1, 2, 4], ["one", "two", "four"], [10, 20, 40]])
        self.assertEqual(call("HLOOKUP", 2, htable, 2, False), "two")
        self.assertEqual(call("HLOOKUP", 3, htable, 3), 20.0)
        self.assertEqual(call("HLOOKUP", 2, htable, 4, False), REF)
        self.assertEqual(call("HLOOKUP", 9, htable, 2, False), NA)

    def test_index(self):
        arr = Area([[1, 2, 3], [4, 5, 6]])
        self.assertEqual(call("INDEX", arr, 2, 3), 6.0)
        self.assertEqual(call("INDEX", arr, 1, 0), Area([[1, 2, 3]]))
        self.assertEqual(call("INDEX", arr, 0, 2), Area([[2], [5]]))
        self.assertEqual(call("INDEX", arr, 0, 0), arr)
        self.assertEqual(call("INDEX", arr, 2), Area([[4, 5, 6]]))
        self.assertEqual(call("INDEX", arr, 3, 1), REF)
        self.assertEqual(call("INDEX", arr, 1, 4), REF)
        self.assertEqual(call("INDEX", arr, -1, 1), VALUE)

    def test_index_single_row_or_column_vector(self):
        self.assertEqual(call("INDEX", Area([[10, 20, 30]]), 2), 20.0)
        self.assertEqual(call("INDEX", Area([[10], [20]]), 2), 20.0)
        self.assertEqual(call("INDEX", Area([[10]]), 1), 10.0)

    def test_xlookup_basics(self):
        self.assertEqual(
            call("XLOOKUP", "b", Area([["a", "b", "c"]]), Area([[1, 2, 3]])), 2.0
        )
        col = Area([[1], [2], [3]])
        self.assertEqual(
            call("XLOOKUP", 9, col, Area([[10], [20], [30]]), "nf"), "nf"
        )
        self.assertEqual(call("XLOOKUP", 9, col, Area([[10], [20], [30]])), NA)
        self.assertEqual(
            call("XLOOKUP", "B", Area([["a", "b"]]), Area([[1, 2]])), 2.0
        )

    def test_xlookup_returns_row_for_2d_return_array(self):
        self.assertEqual(
            call(
                "XLOOKUP",
                2,
                Area([[1], [2], [3]]),
                Area([["x", "y"], ["p", "q"], ["r", "s"]]),
            ),
            Area([["p", "q"]]),
        )

    def test_xlookup_match_modes(self):
        vals = Area([[10], [20], [30]])
        keys = Area([[1], [2], [3]])
        self.assertEqual(call("XLOOKUP", 2.5, keys, vals, "nf", -1), 20.0)
        self.assertEqual(call("XLOOKUP", 2.5, keys, vals, "nf", 1), 30.0)
        self.assertEqual(
            call("XLOOKUP", "b*", Area([["apple"], ["banana"], ["blue"]]),
                 Area([[1], [2], [3]]), "nf", 2),
            2.0,
        )

    def test_xlookup_search_modes(self):
        dup_keys = Area([[1], [3], [3]])
        dup_vals = Area([[10], [20], [30]])
        self.assertEqual(call("XLOOKUP", 3, dup_keys, dup_vals, "nf", 0, 1), 20.0)
        self.assertEqual(call("XLOOKUP", 3, dup_keys, dup_vals, "nf", 0, -1), 30.0)
        sorted_keys = Area([[1], [3], [5], [7]])
        sorted_vals = Area([[10], [30], [50], [70]])
        self.assertEqual(call("XLOOKUP", 5, sorted_keys, sorted_vals, "nf", 0, 2), 50.0)
        self.assertEqual(call("XLOOKUP", 6, sorted_keys, sorted_vals, "nf", 0, 2), "nf")

    def test_xlookup_shape_and_blank_and_errors(self):
        self.assertEqual(call("XLOOKUP", 1, Area([[1], [2]]), Area([[1], [2], [3]])), VALUE)
        self.assertEqual(call("XLOOKUP", 1, Area([[1]]), Area([[None]])), 0.0)
        self.assertEqual(
            call("XLOOKUP", DIV0, Area([[1]]), Area([[2]])), DIV0
        )


class ConditionalFunctionTests(unittest.TestCase):
    def test_sumif_criteria_operators(self):
        self.assertEqual(call("SUMIF", Area([[1], [5], [7]]), ">4"), 12.0)
        self.assertEqual(call("SUMIF", Area([[1], [2], [3]]), ">=2", Area([[10], [20], [30]])), 50.0)
        self.assertEqual(call("SUMIF", Area([[1]]), "<=3", Area([[10]])), 10.0)
        self.assertEqual(call("SUMIF", Area([[0.5], [2]]), "<1"), 0.5)
        self.assertEqual(call("SUMIF", Area([[1], [5]]), Area.scalar(">4")), 5.0)

    def test_sumif_text_and_case_and_wildcards(self):
        self.assertEqual(
            call("SUMIF", Area([["a"], ["b"], ["A"]]), "a", Area([[1], [2], [4]])), 5.0
        )
        self.assertEqual(
            call("SUMIF", Area([["apple"], ["apricot"], ["banana"]]), "ap*",
                 Area([[1], [2], [4]])),
            3.0,
        )
        self.assertEqual(
            call("SUMIF", Area([["b"], ["a"], ["c"]]), ">a", Area([[1], [2], [4]])), 5.0
        )

    def test_sumif_numeric_text_and_not_equal_and_blank(self):
        self.assertEqual(call("SUMIF", Area([[5], ["5"], [6]]), 5), 5.0)
        self.assertEqual(
            call("SUMIF", Area([["x"], [None], [2.0]]), "<>x", Area([[1], [10], [100]])),
            110.0,
        )
        self.assertEqual(
            call("SUMIF", Area([[None], [1.0]]), "=", Area([[7], [9]])), 7.0
        )

    def test_sumif_error_in_matched_sum_cell(self):
        self.assertEqual(call("SUMIF", Area([[1], [2]]), 1, Area([[DIV0], [3]])), DIV0)
        self.assertEqual(call("SUMIF", Area([[1]]), NA), NA)

    def test_sumifs(self):
        sums = Area([[10], [20], [30], [40]])
        r1 = Area([["a"], ["b"], ["a"], ["a"]])
        r2 = Area([[1], [2], [3], [4]])
        self.assertEqual(call("SUMIFS", sums, r1, "a", r2, ">2"), 70.0)
        self.assertEqual(call("SUMIFS", Area([[1], [2]]), Area([[1]]), 1), VALUE)
        self.assertEqual(call("SUMIFS", sums), VALUE)

    def test_countif(self):
        self.assertEqual(call("COUNTIF", Area([[1], [2], [3], ["3"]]), 3), 2.0)
        self.assertEqual(call("COUNTIF", Area([["a"], ["A"], ["ab"]]), "a"), 2.0)
        self.assertEqual(call("COUNTIF", Area([["ab"], ["a"]]), "a?"), 1.0)
        self.assertEqual(call("COUNTIF", Area([["x*"], ["xy"]]), "x~*"), 1.0)
        self.assertEqual(call("COUNTIF", Area([[True], [False], ["x"]]), "TRUE"), 1.0)
        # an empty criteria cell is 0: blank and empty-text cells do not match it, a 0 does
        self.assertEqual(call("COUNTIF", Area([[None], ["x"], [""], [0.0]]), None), 1.0)
        self.assertEqual(call("COUNTIF", Area([[1], [5]]), ">=0"), 2.0)

    def test_countifs(self):
        r1 = Area([["a"], ["b"], ["a"]])
        r2 = Area([[1], [2], [3]])
        self.assertEqual(call("COUNTIFS", r1, "a", r2, ">1"), 1.0)
        self.assertEqual(call("COUNTIFS", r1, "a", r2, ">0"), 2.0)
        self.assertEqual(call("COUNTIFS", r1, "a", Area([[1]]), 1), VALUE)

    def test_maxifs_minifs(self):
        vals = Area([[5], [9], [7]])
        keys = Area([["a"], ["b"], ["a"]])
        self.assertEqual(call("MAXIFS", vals, keys, "a"), 7.0)
        self.assertEqual(call("MINIFS", vals, keys, "a"), 5.0)
        self.assertEqual(call("MAXIFS", Area([[5]]), Area([["b"]]), "a"), 0.0)
        self.assertEqual(call("MINIFS", Area([[5]]), Area([["b"]]), "a"), 0.0)
        self.assertEqual(
            call("MAXIFS", Area([[1], [6], [9]]), Area([[1], [6], [9]]), ">5"), 9.0
        )


class DynamicArrayTests(unittest.TestCase):
    def test_sequence_shapes(self):
        self.assertEqual(call("SEQUENCE", 3), Area([[1], [2], [3]]))
        self.assertEqual(call("SEQUENCE", 2, 3), Area([[1, 2, 3], [4, 5, 6]]))
        self.assertEqual(call("SEQUENCE", 2, 2, 10, 5), Area([[10, 15], [20, 25]]))
        self.assertEqual(call("SEQUENCE", 2.9), Area([[1], [2]]))
        self.assertEqual(call("SEQUENCE", "x"), VALUE)
        self.assertEqual(call("SEQUENCE", NA), NA)
        self.assertEqual(call("SEQUENCE", 0), CALC)

    def test_unique_rows(self):
        self.assertEqual(
            call("UNIQUE", Area([[1], [2], [1], [3]])), Area([[1], [2], [3]])
        )
        self.assertEqual(call("UNIQUE", Area([["a"], ["A"], ["b"]])), Area([["a"], ["b"]]))
        self.assertEqual(
            call("UNIQUE", Area([[1, 2], [1, 2], [3, 4]])), Area([[1, 2], [3, 4]])
        )

    def test_unique_exactly_once_and_by_col(self):
        self.assertEqual(
            call("UNIQUE", Area([[1, 2], [1, 2], [3, 4]]), False, True), Area([[3, 4]])
        )
        self.assertEqual(
            call("UNIQUE", Area([[1, 2, 1], [3, 4, 3]]), True), Area([[1, 2], [3, 4]])
        )
        self.assertEqual(call("UNIQUE", Area([[1], [1]]), False, True), CALC)
        self.assertEqual(call("UNIQUE", 5), Area([[5]]))


class LiftingAndPropagationTests(unittest.TestCase):
    def test_scalar_function_lifts_over_area(self):
        self.assertEqual(
            call("SUBSTITUTE", Area([["aa", "ba"]]), "a", "c"), Area([["cc", "bc"]])
        )
        self.assertEqual(call("LEN", Area([["ab", ""], [None, "c"]])),
                         Area([[2.0, 0.0], [0.0, 1.0]]))
        self.assertEqual(call("TRIM", Area([["  x  ", "y "]])), Area([["x", "y"]]))

    def test_lift_mismatched_shapes_is_na(self):
        self.assertEqual(
            call("SUBSTITUTE", Area([["a", "b"]]), Area([["a"], ["b"]]), "x"), NA
        )

    def test_one_by_one_area_acts_as_scalar(self):
        self.assertEqual(call("SUBSTITUTE", Area([["aa"]]), "a", Area.scalar("b")), "bb")

    def test_lift_keeps_per_cell_errors(self):
        self.assertEqual(call("LEFT", Area([[DIV0, "ab"]]), 1), Area([[DIV0, "a"]]))

    def test_non_propagating_info_functions_lift_without_erroring(self):
        self.assertEqual(call("ISERROR", Area([[NA, 1.0]])), Area([[True, False]]))
        self.assertEqual(call("ISNUMBER", Area([[NA, 1.0]])), Area([[False, True]]))

    def test_error_propagation_across_functions(self):
        self.assertEqual(call("MID", DIV0, 1, 1), DIV0)
        self.assertEqual(call("INT", NA), NA)
        self.assertEqual(call("FLOOR", DIV0, 1), DIV0)
        self.assertEqual(call("MOD", 1, NA), NA)
        self.assertEqual(call("DAYS", 1, DIV0), DIV0)
        self.assertEqual(call("VLOOKUP", DIV0, TABLE, 2, False), DIV0)
        self.assertEqual(call("MATCH", NA, Area([[1, 2]]), 0), NA)

    def test_never_raises_on_bad_input(self):
        self.assertEqual(call("INT", "abc"), VALUE)
        self.assertEqual(call("TEXT", "abc", "0.00"), "abc")
        self.assertEqual(call("SEQUENCE", -3), CALC)


class RegistryTests(unittest.TestCase):
    def test_required_names_present_and_no_xlfn_prefix(self):
        names = [
            "SUBSTITUTE", "TRIM", "NUMBERVALUE", "CONCAT", "CONCATENATE", "LEFT",
            "RIGHT", "MID", "FIND", "SEARCH", "LEN", "VALUE", "TEXT", "INT",
            "FLOOR", "FLOOR.MATH", "ROUNDUP", "MOD", "SUM", "MAX", "MIN",
            "SUMPRODUCT", "ISEVEN", "DAYS", "ISERROR", "ISNUMBER", "ISBLANK",
            "ISTEXT", "NOT", "AND", "OR", "COUNTA", "MATCH", "VLOOKUP", "HLOOKUP",
            "INDEX", "XLOOKUP", "SUMIF", "SUMIFS", "COUNTIF", "COUNTIFS",
            "MAXIFS", "MINIFS", "SEQUENCE", "UNIQUE",
        ]
        for name in names:
            self.assertIn(name, FUNCTIONS)
        self.assertFalse([k for k in FUNCTIONS if k.startswith("_XLFN")])
        self.assertTrue(all(k == k.upper() for k in FUNCTIONS))

    def test_lazy_set(self):
        self.assertEqual(
            LAZY,
            {"IF", "IFERROR", "IFS", "SWITCH", "LET", "COLUMN", "ROW", "ISFORMULA",
             "CELL", "OFFSET", "INDIRECT", "AND", "OR"},
        )
        self.assertIn("AND", FUNCTIONS)
        self.assertIn("OR", FUNCTIONS)
        self.assertNotIn("IF", FUNCTIONS)


class LookupPerformanceTests(unittest.TestCase):
    def test_100k_exact_vlookups_on_5k_rows_under_a_second(self):
        table = Area([[float(i), f"n{i}", float(2 * i)] for i in range(5000)])
        fn = FUNCTIONS["VLOOKUP"]
        start = time.perf_counter()
        total = 0.0
        for i in range(100_000):
            total += fn(float(i % 5000), table, 3.0, False)
        elapsed = time.perf_counter() - start
        expected = 20 * sum(2.0 * k for k in range(5000))
        self.assertEqual(total, expected)
        self.assertTrue(table.memo)
        self.assertLess(elapsed, 1.0)


if __name__ == "__main__":
    unittest.main()
