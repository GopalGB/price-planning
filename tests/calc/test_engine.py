import unittest

from pp.calc.engine import Book, Shared
from pp.calc.model import CELL, Template, column_number, relative_key
from pp.calc.values import XlError


def workbook(sheets: dict, arrays=()) -> Template:
    """A template from {"Sheet": {"A1": value or "=formula"}}; formulas start with Excel's result unknown."""
    names = list(sheets)
    formulas, families, keyed, constants, texts = {}, [], {}, {}, {}
    for index, name in enumerate(names):
        for address, value in sheets[name].items():
            match = CELL.fullmatch(address)
            key = (index, int(match.group(4)) - 1, column_number(match.group(2)) - 1)
            if isinstance(value, str) and value.startswith("="):
                array = f"{name}!{address}" in arrays
                family_key = (index, array, relative_key(value, key[1], key[2]))
                if family_key not in keyed:
                    keyed[family_key] = len(families)
                    families.append((key, value, array))
                formulas[key] = keyed[family_key]
                texts[key] = value
            else:
                constants[key] = float(value) if isinstance(value, int) else value
    return Template(names, formulas, families, {}, constants, texts)


def evaluate(sheets, cell="Sheet1!A1", arrays=()):
    template = workbook(sheets, arrays)
    book = Book(Shared(template))
    name, address = cell.split("!")
    match = CELL.fullmatch(address)
    key = (template.names.index(name), int(match.group(4)) - 1, column_number(match.group(2)) - 1)
    return book.recalculate(list(template.formulas))[key], book


class FamilyTests(unittest.TestCase):
    def test_copied_formulas_share_a_family_and_absolute_parts_stay_fixed(self):
        template = workbook({"Sheet1": {"A1": 1, "A2": 2, "B1": "=A1*$C$1", "B2": "=A2*$C$1", "C1": 10}})
        self.assertEqual(len(template.families), 1)
        self.assertNotEqual(relative_key("=A1+'My Sheet'!B2", 0, 0), relative_key("=A1+'My Sheet'!B3", 0, 0))


class EvaluationTests(unittest.TestCase):
    def test_arithmetic_text_and_comparison(self):
        value, _ = evaluate({"Sheet1": {"A1": '=B1*2+C1&"x"', "B1": 3, "C1": 1}})
        self.assertEqual(value, "7x")
        self.assertIs(evaluate({"Sheet1": {"A1": '=B1="abc"', "B1": "ABC"}})[0], True)

    def test_blank_reference_reads_as_zero_and_errors_propagate(self):
        self.assertEqual(evaluate({"Sheet1": {"A1": "=B1"}})[0], 0.0)
        self.assertEqual(evaluate({"Sheet1": {"A1": "=1/B1"}})[0], XlError("#DIV/0!"))
        self.assertEqual(evaluate({"Sheet1": {"A1": "=IFERROR(1/B1,5)"}})[0], 5.0)

    def test_if_evaluates_only_the_chosen_branch(self):
        self.assertEqual(evaluate({"Sheet1": {"A1": '=IF(B1>1,"big",1/0)', "B1": 5}})[0], "big")

    def test_chains_settle_on_demand_across_sheets(self):
        value, _ = evaluate({"Sheet1": {"A1": "=Other!A1+1"}, "Other": {"A1": "=B1*10", "B1": 4}})
        self.assertEqual(value, 41.0)

    def test_let_binds_names(self):
        self.assertEqual(evaluate({"Sheet1": {"A1": "=_xlfn.LET(_xlpm.x, B1*2, _xlpm.x+_xlpm.x)", "B1": 3}})[0], 12.0)

    def test_lookup_over_a_range_and_column(self):
        cells = {"A1": "=VLOOKUP(\"b\",D1:E3,2,FALSE)", "B1": "=COLUMN(E1)", "D1": "a", "E1": 1, "D2": "b",
                 "E2": 2, "D3": "c", "E3": 3}
        self.assertEqual(evaluate({"Sheet1": cells})[0], 2.0)
        self.assertEqual(evaluate({"Sheet1": cells}, "Sheet1!B1")[0], 5.0)

    def test_non_array_formula_intersects_a_column_range_with_its_own_row(self):
        cells = {"A2": "=B1:B3*2", "B1": 1, "B2": 7, "B3": 3}
        self.assertEqual(evaluate({"Sheet1": cells}, "Sheet1!A2")[0], 14.0)

    def test_array_formula_evaluates_arrays_element_wise(self):
        cells = {"A1": '=MATCH("b2",D1:D3&E1:E3,0)', "D1": "a", "E1": 1, "D2": "b", "E2": 2, "D3": "c", "E3": 3}
        self.assertEqual(evaluate({"Sheet1": cells}, arrays={"Sheet1!A1"})[0], 2.0)

    def test_sumproduct_takes_arrays_in_a_plain_formula(self):
        cells = {"A1": "=SUMPRODUCT((B1:B3>1)*C1:C3)", "B1": 1, "B2": 2, "B3": 3, "C1": 10, "C2": 20, "C3": 30}
        self.assertEqual(evaluate({"Sheet1": cells})[0], 50.0)

    def test_an_edit_recalculates_only_what_is_stale(self):
        template = workbook({"Sheet1": {"A1": "=B1+1", "B1": 1}})
        template.results[(0, 0, 0)] = 2.0
        book = Book(Shared(template))
        self.assertEqual(book.cell(0, 0, 0), 2.0)
        book.inputs[(0, 0, 1)] = 5.0
        self.assertEqual(book.recalculate([(0, 0, 0)]), {(0, 0, 0): 6.0})
        self.assertEqual(book.results, {(0, 0, 0): 6.0})

    def test_whole_row_absolute_reference_stays_on_its_row(self):
        cells = {"A5": '=MATCH("q",$2:$2,0)', "C2": "q", "C5": "q"}
        self.assertEqual(evaluate({"Sheet1": cells}, "Sheet1!A5")[0], 3.0)

    def test_omitted_xlookup_arguments_take_their_defaults(self):
        cells = {"A1": '=IFERROR(_xlfn.XLOOKUP("zz",C1:C2,D1:D2,,,1),"-")', "C1": "a", "C2": "b", "D1": 1, "D2": 2}
        self.assertEqual(evaluate({"Sheet1": cells})[0], "-")

    def test_switch_does_not_match_an_empty_text_to_a_blank_case_cell(self):
        cells = {"A1": '=_xlfn.SWITCH(IF(B1=0,"",B1),C1,"Yellow","")'}
        self.assertEqual(evaluate({"Sheet1": cells})[0], "")

    def test_a_blank_lookup_key_looks_up_zero_and_skips_blank_rows(self):
        cells = {"A1": "=VLOOKUP(B1,C1:D3,2,0)", "C1": "x", "D1": "blank row", "C3": 0, "D3": "zero row"}
        cells["D2"] = "empty key row"
        self.assertEqual(evaluate({"Sheet1": cells})[0], "zero row")

    def test_rank_counts_only_numbers_and_takes_an_order(self):
        cells = {"A1": "=RANK(B2,B1:B4)", "A2": "=RANK(B2,B1:B4,1)", "A3": "=RANK(9,B1:B4)",
                 "B1": 5, "B2": 3, "B3": "7", "B4": 1}
        self.assertEqual(evaluate({"Sheet1": cells})[0], 2.0)
        self.assertEqual(evaluate({"Sheet1": cells}, "Sheet1!A2")[0], 2.0)
        self.assertEqual(evaluate({"Sheet1": cells}, "Sheet1!A3")[0], XlError("#N/A"))

    def test_unknown_function_keeps_excels_result(self):
        template = workbook({"Sheet1": {"A1": "=NOSUCHFN(1)"}})
        template.results[(0, 0, 0)] = 9.0
        book = Book(Shared(template))
        self.assertEqual(book.recalculate([(0, 0, 0)]), {(0, 0, 0): 9.0})
        self.assertEqual(len(book.shared.unsupported), 1)


if __name__ == "__main__":
    unittest.main()
