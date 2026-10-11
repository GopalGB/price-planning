import importlib.util
import json
import unittest

from ops.fixtures import make_dashboard
from tests.v2.support import BackendFixture, QueryBoundaryDatabase


def small_dashboard(path):
    make_dashboard(path, row_count=240)


def tiny_dashboard(path):
    make_dashboard(path, row_count=5)


def options_all(pivot, slicers):
    return {
        "page": [{"field": item["field"], "items": None} for item in pivot["page_fields"]],
        "slicers": [{"name": item["name"], "field": item["field"], "items": None}
                    for item in slicers],
        "collapsed": [],
    }


class PivotServiceTests(unittest.TestCase):
    def setUp(self):
        self.fixture = BackendFixture(small_dashboard)
        self.pivot = self.fixture.snapshot.sheets[0]["pivots"][0]
        self.slicers = self.fixture.snapshot.sheets[0]["slicers"]

    def tearDown(self):
        self.fixture.close()

    def service(self):
        from pp.pivot import PivotService

        return PivotService(self.fixture.database, self.fixture.snapshot)

    def refresh(self, options=None, user="owner"):
        from pp.pivot import PivotQuery

        return self.service().refresh(PivotQuery("copy-a", user, 0, "DashboardPivot", options or {}))

    def test_pivot_service_module_exists(self):
        self.assertIsNotNone(importlib.util.find_spec("pp.pivot"), "pp.pivot service is missing")

    def test_read_requires_access_and_current_calculation_revision(self):
        from pp.errors import AppError

        with self.assertRaises(AppError) as denied:
            self.refresh(user="outsider")
        self.assertEqual((denied.exception.status, denied.exception.code), (404, "not_found"))

        with self.fixture.database.transaction() as connection:
            connection.execute("UPDATE pp_copies SET input_revision=1 WHERE id='copy-a'")
        with self.assertRaises(AppError) as pending:
            self.refresh()
        self.assertEqual((pending.exception.status, pending.exception.code), (409, "calculation_pending"))
        self.assertEqual(pending.exception.extra, {"input_revision": 1, "calculated_revision": 0})

    def test_cached_refresh_reauthorizes_after_access_is_revoked(self):
        from pp.errors import AppError

        service = self.service()
        query = self.query(options_all(self.pivot, self.slicers), user="viewer")
        service.refresh(query)
        with self.fixture.database.transaction() as connection:
            connection.execute("DELETE FROM pp_copy_shares WHERE copy_id='copy-a' AND user_id='viewer'")

        with self.assertRaises(AppError) as denied:
            service.refresh(query)
        self.assertEqual((denied.exception.status, denied.exception.code), (404, "not_found"))

    def test_same_field_filters_intersect_and_hidden_items_stay_hidden(self):
        options = options_all(self.pivot, self.slicers)
        options["slicers"] = [
            {"name": "Slicer_Region", "field": 0, "items": [0, 1]},
            {"name": "second-region-view", "field": 0, "items": [1, 3]},
            *options["slicers"][1:],
        ]

        result = self.refresh(options)
        item_rows = [row for row in result["rows"] if row["kind"] == "item"]

        self.assertEqual([row["caption"] for row in item_rows], ["Region 02"])
        self.assertNotIn("Region 04", [row["caption"] for row in result["rows"]])
        self.assertEqual(result["filters"]["slicers"][0]["items"], [0, 1])

    def test_manual_and_data_sort_preserve_excel_order_rules(self):
        field = self.pivot["fields"][0]
        field["sort"] = {"type": "manual", "rank_by": None, "data_source": False}

        manual = self.refresh(options_all(self.pivot, self.slicers))
        manual_rows = [row["caption"] for row in manual["rows"] if row["kind"] == "item"]
        self.assertEqual(manual_rows[:3], ["Region 03", "Region 01", "Region 02"])
        self.assertEqual(manual_rows[-1], "(blank)")

        field["sort"] = {"type": "descending", "rank_by": 6, "data_source": True}
        ranked = self.service().refresh(self.query(options_all(self.pivot, self.slicers)))
        revenue_by_row = {
            cell["r"]: cell["v"] for cell in ranked["cells"] if cell.get("data_field") == 6
        }
        data_rows = [row for row in ranked["rows"] if row["kind"] == "item"]
        values = [revenue_by_row[row["r"]] for row in data_rows if row["caption"] != "(blank)"]
        self.assertEqual(values, sorted(values, reverse=True))
        self.assertEqual(data_rows[-1]["caption"], "(blank)")

    def test_invalid_or_nonfinite_filter_selectors_are_rejected(self):
        from pp.errors import AppError

        options = options_all(self.pivot, self.slicers)
        for value in (9999, {"kind": "number", "value": float("inf")},
                      {"kind": "boolean", "value": 1}):
            with self.subTest(value=value):
                options["slicers"][0]["items"] = [value]
                with self.assertRaises(AppError) as error:
                    self.refresh(options)
                self.assertEqual((error.exception.status, error.exception.code), (422, "invalid"))

    def query(self, options, user="owner"):
        from pp.pivot import PivotQuery

        return PivotQuery("copy-a", user, 0, "DashboardPivot", options)

    def test_outline_rows_subtotals_and_collapse_are_structural(self):
        self.pivot["row_fields"] = [0, 1]
        department = self.pivot["fields"][1]
        department["axis"] = "row"
        department["subtotals"] = ["sum"]
        expanded = self.refresh(options_all(self.pivot, self.slicers))
        parent = next(row for row in expanded["rows"] if row["kind"] == "subtotal")
        self.assertTrue(parent["collapsible"])

        options = options_all(self.pivot, self.slicers)
        options["collapsed"] = [{"field": parent["field"], "item": parent["item"]}]
        collapsed = self.refresh(options)
        collapsed_parent = next(row for row in collapsed["rows"] if row["caption"] == parent["caption"])
        self.assertTrue(collapsed_parent["collapsed"])
        self.assertLess(len(collapsed["rows"]), len(expanded["rows"]))
        self.assertEqual(collapsed["rows"][-1]["kind"], "grand_total")

        self.pivot["row_fields"] = [0]
        flat = self.service().refresh(self.query(options_all(self.pivot, self.slicers)))
        self.assertFalse(any(row["collapsible"] for row in flat["rows"]))

    def test_column_fields_include_values_and_column_grand_totals(self):
        self.pivot["column_fields"] = [4, -2]
        self.pivot["fields"][4]["axis"] = "column"
        result = self.refresh(options_all(self.pivot, self.slicers))
        revenue = [column for column in result["columns"] if column["data_field"] == 6]
        grand_column = next(column for column in revenue if column["grand_total"])
        class_columns = [column for column in revenue if not column["grand_total"]]
        grand_row = result["rows"][-1]["r"]
        by_column = {cell["c"]: cell["v"] for cell in result["cells"]
                     if cell["r"] == grand_row and cell.get("data_field") == 6}

        self.assertTrue(class_columns)
        self.assertTrue(all(column["path"] for column in class_columns))
        self.assertEqual(sum(by_column[column["c"]] for column in class_columns),
                         by_column[grand_column["c"]])

    def test_revision_change_invalidates_result_cache_and_reads_current_result(self):
        options = options_all(self.pivot, self.slicers)
        service = self.service()
        first = service.refresh(self.query(options))
        first_price = next(cell["v"] for cell in first["cells"]
                           if cell["row_kind"] == "grand_total" and cell.get("data_field") == 9)
        with self.fixture.database.transaction() as connection:
            connection.execute(
                "INSERT INTO pp_results(copy_id,sheet,row,col,value_json,calculated_revision) VALUES(?,?,?,?,?,?)",
                ("copy-a", 1, 2, 9, json.dumps(999), 1),
            )
            connection.execute(
                "UPDATE pp_copies SET input_revision=1,calculated_revision=1 WHERE id='copy-a'",
            )
        second = service.refresh(self.query(options))
        second_price = next(cell["v"] for cell in second["cells"]
                            if cell["row_kind"] == "grand_total" and cell.get("data_field") == 9)

        self.assertNotEqual(second_price, first_price)
        self.assertEqual(second_price - first_price, 996)
        self.assertLessEqual(service.cache_info()["results"], 32)

    def test_revision_and_overlays_are_read_from_one_database_snapshot(self):
        from pp.pivot import PivotService

        with self.fixture.database.transaction() as connection:
            connection.execute(
                "INSERT INTO pp_results(copy_id,sheet,row,col,value_json,calculated_revision) "
                "VALUES(?,?,?,?,?,?)", ("copy-a", 1, 2, 9, json.dumps(100), 0),
            )

        def advance_revision():
            with self.fixture.database.transaction() as connection:
                connection.execute(
                    "UPDATE pp_results SET value_json=?,calculated_revision=1 "
                    "WHERE copy_id='copy-a' AND sheet=1 AND row=2 AND col=9", (json.dumps(999),),
                )
                connection.execute(
                    "UPDATE pp_copies SET input_revision=1,calculated_revision=1 WHERE id='copy-a'",
                )

        database = QueryBoundaryDatabase(self.fixture.database, "SELECT row,col,value_json FROM pp_overrides",
                                         advance_revision)
        service = PivotService(database, self.fixture.snapshot)
        options = options_all(self.pivot, self.slicers)
        first = service.refresh(self.query(options))
        second = service.refresh(self.query(options))
        totals = [next(cell["v"] for cell in result["cells"]
                       if cell["row_kind"] == "grand_total" and cell.get("data_field") == 9)
                  for result in (first, second)]

        self.assertTrue(database.triggered)
        self.assertEqual(first["pivot"]["calculated_revision"], 0)
        self.assertEqual(second["pivot"]["calculated_revision"], 1)
        self.assertEqual(totals[1] - totals[0], 899)

    def test_current_dynamic_axis_values_have_stable_selectors(self):
        with self.fixture.database.transaction() as connection:
            connection.executemany(
                "INSERT INTO pp_results(copy_id,sheet,row,col,value_json,calculated_revision) "
                "VALUES(?,?,?,?,?,?)",
                [("copy-a", 1, 1, 0, json.dumps("Region 99"), 0),
                 ("copy-a", 1, 1, 1, json.dumps("Department 99"), 0)],
            )
        options = options_all(self.pivot, self.slicers)
        result = self.refresh(options)
        dynamic = next(item for item in result["slicers"][0]["items"]
                       if item["caption"] == "Region 99")

        self.assertIn("Region 99", [row["caption"] for row in result["rows"]])
        self.assertIsNone(dynamic["cache_index"])
        self.assertEqual(dynamic["selector"], {"kind": "string", "value": "Region 99"})
        self.assertTrue(dynamic["available"])

        selector = dynamic["selector"]
        options["slicers"][0]["items"] = [selector]
        options["page"][0]["items"] = [{"kind": "string", "value": "Department 99"}]
        selected = self.refresh(options)
        self.assertEqual([row["caption"] for row in selected["rows"] if row["kind"] == "item"],
                         ["Region 99"])


class PivotValueHandlingTests(unittest.TestCase):
    def setUp(self):
        self.fixture = BackendFixture(tiny_dashboard)
        self.pivot = self.fixture.snapshot.sheets[0]["pivots"][0]
        self.slicers = self.fixture.snapshot.sheets[0]["slicers"]

    def tearDown(self):
        self.fixture.close()

    def test_sum_propagates_errors_and_ignores_other_nonnumeric_values(self):
        from pp.pivot import PivotQuery, PivotService

        values = [{"error": "#DIV/0!"}, "text", None, True, 7]
        with self.fixture.database.transaction() as connection:
            connection.executemany(
                "INSERT INTO pp_results(copy_id,sheet,row,col,value_json,calculated_revision) VALUES(?,?,?,?,?,?)",
                [("copy-a", 1, row, 9, json.dumps(value), 0) for row, value in enumerate(values, 1)],
            )
        options = options_all(self.pivot, self.slicers)
        result = PivotService(self.fixture.database, self.fixture.snapshot).refresh(
            PivotQuery("copy-a", "owner", 0, "DashboardPivot", options)
        )
        total = next(cell for cell in result["cells"]
                     if cell["row_kind"] == "grand_total" and cell.get("data_field") == 9)

        self.assertEqual(total["v"], {"error": "#DIV/0!"})
        self.assertEqual(total["number_format"], "General")
        self.assertTrue(total["locked"])

        with self.fixture.database.transaction() as connection:
            connection.execute(
                "UPDATE pp_results SET value_json=? WHERE copy_id=? AND sheet=1 AND row=1 AND col=9",
                (json.dumps(None), "copy-a"),
            )
        result = PivotService(self.fixture.database, self.fixture.snapshot).refresh(
            PivotQuery("copy-a", "owner", 0, "DashboardPivot", options)
        )
        total = next(cell for cell in result["cells"]
                     if cell["row_kind"] == "grand_total" and cell.get("data_field") == 9)
        self.assertEqual(total["v"], 7)


if __name__ == "__main__":
    unittest.main()
