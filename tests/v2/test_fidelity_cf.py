"""Behavioral conditional-formatting fidelity coverage."""

from __future__ import annotations

import json
import tempfile
import unittest
from unittest.mock import patch
from pathlib import Path
from types import SimpleNamespace
import xml.etree.ElementTree as ET

from ops.fixtures import (
    make_broad_conditional_formatting,
    make_conditional_formatting,
    make_volatile_conditional_formatting,
)
from pp.calc_service import CalculationPublication, CalculationService
from pp.engine.formualizer import EngineError, FormualizerEngine
from pp.fidelity.conditional_formatting import (
    MAX_FORMULA_TARGETS,
    MAX_OVERRIDE_RECORDS,
    CurrentCopyValues,
    conditional_style_overrides,
)
from pp.template.importer import CF_BORDER_STYLES, NS, _dxf_style, import_template
from pp.workbook_service import WorkbookService
from tests.v2.api_support import APIFixture


class ConditionalFormattingImportTest(unittest.TestCase):
    def test_imports_ooxml_rules_and_literal_differential_styles(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            source = Path(directory) / "synthetic-cf.xlsm"
            target = Path(directory) / "template"
            make_conditional_formatting(source)
            snapshot = import_template(source, target)
            imported = json.loads((snapshot / "meta.json").read_text())
            metadata = imported["sheets"][0]

        rules = metadata["conditional_formats"]
        self.assertEqual(
            [(rule["type"], rule["range"], rule["dxf"]) for rule in rules[:6]],
            [
                ("expression", [1, 0, 2, 0], {"fill": "#FF0000"}),
                ("cellIs", [1, 1, 1, 1], {"color": "#00AA00"}),
                ("containsText", [1, 2, 1, 2], {"bold": True}),
                ("duplicateValues", [1, 3, 1025, 3], {"border": {"b": {"s": 1, "cl": {"rgb": "#0000FF"}}}}),
                ("beginsWith", [1, 4, 1, 4], {"nf": "0.000"}),
                ("containsErrors", [1, 5, 1, 6], {"bold": False, "fill": None}),
            ],
        )
        self.assertEqual(rules[20]["dxf"], {"color": "#000000"})
        self.assertEqual(rules[21]["dxf"], {"fill": "#FF0000"})
        self.assertEqual(
            rules[22]["dxf"],
            {"border": {"t": {"s": 1, "cl": {"rgb": "#808080"}}}},
        )
        self.assertEqual(
            [(rule["range"], rule["top_left"]) for rule in rules[26:28]],
            [([1, 26, 2, 26], [1, 26]), ([4, 28, 5, 28], [1, 26])],
        )
        self.assertEqual(rules[29]["dxf"], {})
        self.assertEqual(rules[30]["range"], [0, 16_383, 1_048_575, 16_383])
        self.assertEqual(
            [sheet["has_conditional_formats"] for sheet in imported["sheets"]],
            [True, True, True, True, True, False],
        )

    def test_dxf_border_styles_use_verified_univer_enums_and_skip_empty_sides(self) -> None:
        expected = {
            "thin": 1, "hair": 2, "dotted": 3, "dashed": 4,
            "dashDot": 5, "dashDotDot": 6, "double": 7, "medium": 8,
            "mediumDashed": 9, "mediumDashDot": 10,
            "mediumDashDotDot": 11, "slantDashDot": 12, "thick": 13,
        }
        self.assertEqual(CF_BORDER_STYLES, expected)
        namespace = NS[1:-1]
        for name, value in expected.items():
            node = ET.fromstring(
                f'<dxf xmlns="{namespace}"><border><top style="{name}"/></border></dxf>',
            )
            self.assertEqual(_dxf_style(node, {})["border"]["t"]["s"], value)
        empty = ET.fromstring(
            f'<dxf xmlns="{namespace}"><border><top/><right style="none"/>'
            '<bottom style="unknown"/></border></dxf>',
        )
        self.assertNotIn("border", _dxf_style(empty, {}))


class ConditionalFormattingRowsTest(unittest.TestCase):
    def setUp(self) -> None:
        self.fixture = APIFixture(make_conditional_formatting)
        self.client, self.csrf, user = self.fixture.signup("fidelity@example.test")
        _, copy = self.fixture.make_copy(self.client, self.csrf)
        self.copy_id = copy["id"]
        self.user_id = user["id"]

    def tearDown(self) -> None:
        self.fixture.close()

    def test_rows_use_current_copy_values_across_sheets_and_sparse_blocks(self) -> None:
        services = self.fixture.app.state.services
        committed = services.operations.apply(
            self.copy_id,
            self.user_id,
            {"op_id": "cf-edit", "action": "edit",
             "cells": [{"sheet": 0, "r": 1, "c": 1, "v": 12, "expected_version": 0}]},
        )
        self.assertEqual(committed["input_revision"], 1)
        published = CalculationService(self.fixture.database).complete(
            CalculationPublication(self.copy_id, 1, [
                (0, 1025, 16, 1.0), (0, 1499, 23, 5),
                (1, 1, 1, 24), (1, 2, 1, 8),
            ]))
        self.assertTrue(published)

        self.fixture.snapshot.block.cache_clear()
        first = self.client.get(f"/api/copies/{self.copy_id}/sheets/0/rows?start=1&count=256")
        self.assertEqual(first.status_code, 200)
        body = first.json()
        self.assertEqual(body["conditional_format_warnings"], [])
        self.assertEqual(body["rows"]["1"]["0"]["style_override"], {"fill": "#FF0000"})
        self.assertEqual(body["rows"]["1"]["1"]["style_override"], {"color": "#00AA00"})
        self.assertEqual(body["rows"]["1"]["2"]["style_override"], {"bold": True})
        self.assertEqual(body["rows"]["1"]["3"]["style_override"], {"border": {"b": {"s": 1, "cl": {"rgb": "#0000FF"}}}})
        self.assertEqual(body["rows"]["1"]["4"]["style_override"], {"nf": "0.000"})
        self.assertEqual(body["rows"]["1"]["5"]["style_override"], {"bold": False, "fill": None})
        self.assertNotIn("style_override", body["rows"]["1"]["6"])
        self.assertEqual(body["rows"]["1"]["7"]["style_override"], {"fill": "#111111"})
        self.assertEqual(body["rows"]["1"]["8"]["style_override"], {"fill": "#FF0000"})
        self.assertEqual(body["rows"]["1"]["9"]["style_override"], {"color": "#00AA00"})
        self.assertEqual(body["rows"]["1"]["10"]["style_override"], {"bold": True})
        self.assertEqual(body["rows"]["1"]["11"]["style_override"], {"border": {"b": {"s": 1, "cl": {"rgb": "#0000FF"}}}})
        self.assertEqual(body["rows"]["1"]["12"]["style_override"], {"nf": "0.000"})
        self.assertNotIn("13", body["rows"]["1"])
        self.assertNotIn("14", body["rows"]["1"])
        self.assertNotIn("style_override", body["rows"]["2"]["13"])
        self.assertNotIn("style_override", body["rows"]["2"]["14"])
        self.assertEqual(body["rows"]["1"]["15"]["style_override"], {"fill": "#FF0000"})
        self.assertEqual(body["rows"]["1"]["16"]["style_override"], {"color": "#00AA00"})
        self.assertEqual(body["rows"]["1"]["17"]["style_override"], {"bold": True})
        self.assertEqual(body["rows"]["1"]["18"]["style_override"], {"color": "#000000"})
        self.assertEqual(body["rows"]["1"]["19"]["style_override"], {"fill": "#FF0000"})
        self.assertEqual(
            body["rows"]["1"]["20"]["style_override"],
            {"border": {"t": {"s": 1, "cl": {"rgb": "#808080"}}}},
        )
        combined = {"fill": "#FF0000", "color": "#00AA00"}
        self.assertEqual(body["rows"]["1"]["21"]["style_override"], combined)
        self.assertEqual(body["rows"]["256"]["21"]["style_override"], combined)
        self.assertEqual(body["rows"]["1"]["22"]["style_override"], {"bold": True})
        self.assertEqual(body["rows"]["4"]["28"]["style_override"], {"fill": "#FF0000"})
        self.assertEqual(body["rows"]["1"]["29"]["style_override"], {"bold": True})
        self.assertEqual(body["rows"]["1"]["30"]["style_override"], {})
        base_style = self.fixture.snapshot.styles[body["rows"]["1"]["30"]["s"]]
        self.assertEqual(base_style["fill"], "#445566")
        self.assertNotIn("16383", body["rows"]["1"])
        self.assertNotIn("0", body["rows"].get("2", {}))
        self.assertNotIn("1025", body["rows"])
        self.assertLessEqual(self.fixture.snapshot.block.cache_info().currsize, 2)
        with self.fixture.database.read_transaction() as connection:
            reader = CurrentCopyValues(connection, self.fixture.snapshot, self.copy_id)
            for row in (1, 257, 513, 769, 1025):
                reader.get(0, row, 3)
            self.assertLessEqual(len(reader.blocks), 2)

        scoped = self.client.get(
            f"/api/copies/{self.copy_id}/sheets/1/rows?start=1&count=2",
        ).json()
        self.assertIn("2", scoped["rows"].get("1", {}), scoped)
        self.assertEqual(scoped["rows"]["1"]["2"]["style_override"], {"fill": "#FF0000"})
        global_name = self.client.get(
            f"/api/copies/{self.copy_id}/sheets/2/rows?start=0&count=1",
        ).json()
        self.assertEqual(global_name["rows"]["0"]["1"]["style_override"], {"bold": True})
        for sheet, expected in ((3, {"fill": "#FF0000"}), (4, {"color": "#00AA00"})):
            scratch = self.client.get(
                f"/api/copies/{self.copy_id}/sheets/{sheet}/rows?start=1&count=1",
            ).json()
            self.assertEqual(scratch["rows"]["1"]["0"]["style_override"], expected)

        fork = self.client.post(
            f"/api/copies/{self.copy_id}/fork", {"name": "CF Fork"}, {"x-csrf-token": self.csrf},
        ).json()["copy"]
        fork_published = CalculationService(self.fixture.database).complete(
            CalculationPublication(fork["id"], 0, [
                (0, 1025, 16, 1.0), (0, 1499, 23, 5),
                (1, 1, 1, 24), (1, 2, 1, 8),
            ]))
        if not fork_published:
            with self.fixture.database.connect() as connection:
                fork_copy = connection.execute(
                    "SELECT calculated_revision FROM pp_copies WHERE id=?", (fork["id"],),
                ).fetchone()
            self.assertGreaterEqual(fork_copy["calculated_revision"], 0)
        inherited = self.client.get(f"/api/copies/{fork['id']}/sheets/0/rows?start=1&count=2")
        self.assertEqual(inherited.status_code, 200)
        self.assertEqual(inherited.json()["rows"]["1"]["0"]["style_override"], {"fill": "#FF0000"})

        changed = services.operations.apply(
            self.copy_id,
            self.user_id,
            {"op_id": "cf-clear", "action": "edit",
             "cells": [{"sheet": 0, "r": 1, "c": 1, "v": 4, "expected_version": 1}]},
        )
        self.assertEqual(changed["input_revision"], 2)
        republished = CalculationService(self.fixture.database).complete(
            CalculationPublication(self.copy_id, 2, [
                (0, 1025, 16, 1.0), (0, 1499, 23, 5),
                (1, 1, 1, 8), (1, 2, 1, 8),
            ]))
        self.assertTrue(republished)
        after = self.client.get(f"/api/copies/{self.copy_id}/sheets/0/rows?start=1&count=2")
        self.assertEqual(after.status_code, 200)
        self.assertNotIn("0", after.json()["rows"].get("1", {}))

        outside = self.client.get(
            f"/api/copies/{self.copy_id}/sheets/0/rows?start=1999&count=1",
        )
        self.assertEqual(outside.status_code, 200)
        self.assertEqual(outside.json()["rows"], {})

    def test_rule_limits_warn_without_hiding_rows_or_healthy_styles(self) -> None:
        self.fixture.snapshot.sheets[0]["rows"] = 10_001
        with patch("pp.fidelity.conditional_formatting.MAX_DUPLICATE_SCAN_CELLS", 1_000):
            response = self.client.get(
                f"/api/copies/{self.copy_id}/sheets/0/rows?start=1&count=1",
            )
        self.assertEqual(response.status_code, 200)
        body = response.json()
        self.assertEqual(body["rows"]["1"]["2"]["style_override"], {"bold": True})
        self.assertNotIn("style_override", body["rows"]["1"]["15"])
        codes = {warning["code"] for warning in body["conditional_format_warnings"]}
        self.assertIn("formula_reference_limit", codes)
        self.assertIn("duplicate_scan_limit", codes)
        for warning in body["conditional_format_warnings"]:
            self.assertEqual(set(warning), {"code", "message", "rule_id"})
            self.assertNotIn("COUNTIF", warning["message"])

    def test_formula_target_budget_is_total_and_uses_one_public_engine_load(self) -> None:
        original = FormualizerEngine.load
        with patch(
            "pp.fidelity.conditional_formatting.FormualizerEngine.load",
            side_effect=original,
        ) as load:
            response = self.client.get(
                f"/api/copies/{self.copy_id}/sheets/0/rows?start=1&count=2",
            )
        self.assertEqual(response.status_code, 200)
        self.assertEqual(load.call_count, 1)

        with patch("pp.fidelity.conditional_formatting.MAX_FORMULA_TARGETS", 1):
            limited = self.client.get(
                f"/api/copies/{self.copy_id}/sheets/0/rows?start=1&count=2",
            )
        self.assertEqual(limited.status_code, 200)
        body = limited.json()
        self.assertEqual(body["rows"]["1"]["2"]["style_override"], {"bold": True})
        self.assertNotIn("style_override", body["rows"]["1"].get("0", {}))
        self.assertIn(
            "formula_target_limit",
            {warning["code"] for warning in body["conditional_format_warnings"]},
        )

    def test_disjoint_duplicate_rule_does_not_read_candidate_cells(self) -> None:
        class Values:
            calls = 0

            @staticmethod
            def extent(_sheet):
                return 10, 10

            def get(self, *_coordinate):
                self.calls += 1
                return "duplicate"

        rule = {"type": "duplicateValues", "range": [1, 0, 8, 0],
                "priority": 1, "stop_if_true": False, "dxf": {"bold": True}}
        snapshot = SimpleNamespace(sheets=[{"conditional_formats": [rule]}])
        values = Values()
        query = SimpleNamespace(sheet=0, start=256, count=256)
        result = conditional_style_overrides(snapshot, values, query)
        self.assertEqual((result.styles, result.warnings), ({}, []))
        self.assertEqual(values.calls, 0)

        rule["range"] = [1, 0, 1_048_575, 0]
        values = Values()
        visible = SimpleNamespace(sheet=0, start=1, count=1)
        result = conditional_style_overrides(snapshot, values, visible)
        self.assertEqual(result.styles, {(1, 0): {"bold": True}})
        self.assertLessEqual(values.calls, 11)

        rule["range"] = [1, 0, 2, 0]
        with patch("pp.fidelity.conditional_formatting.MAX_DUPLICATE_SCAN_CELLS", 1):
            result = conditional_style_overrides(
                snapshot, Values(), SimpleNamespace(sheet=0, start=1, count=2),
            )
        self.assertEqual(result.styles, {})
        self.assertEqual(result.warnings[0]["code"], "duplicate_scan_limit")

    def test_failed_rule_rolls_back_styles_and_stop_state(self) -> None:
        rules = [
            {"type": "containsText", "range": [0, 0, 0, 0], "priority": 1,
             "stop_if_true": False, "text": "x", "dxf": {"bold": True}},
            {"type": "containsText", "range": [0, 1, 1, 1], "priority": 2,
             "stop_if_true": True, "text": "x", "dxf": {"fill": "#FF0000"}},
            {"type": "containsText", "range": [0, 1, 0, 1], "priority": 3,
             "stop_if_true": False, "text": "x", "dxf": {"color": "#00AA00"}},
        ]
        snapshot = SimpleNamespace(
            sheets=[{"conditional_formats": rules}], meta={"defined_names": []},
        )
        values = SimpleNamespace(extent=lambda _sheet: (2, 2), get=lambda *_args: "x")
        with patch("pp.fidelity.conditional_formatting.MAX_OVERRIDE_RECORDS", 2):
            result = conditional_style_overrides(
                snapshot, values, SimpleNamespace(sheet=0, start=0, count=2),
            )
        self.assertEqual(
            result.styles,
            {(0, 0): {"bold": True}, (0, 1): {"color": "#00AA00"}},
        )
        self.assertEqual(result.warnings[0]["code"], "override_limit")

    def test_row_style_overrides_are_grouped_once(self) -> None:
        class SinglePassStyles(dict):
            passes = 0

            def __iter__(self):
                raise AssertionError("style coordinates were not grouped")

            def items(self):
                self.passes += 1
                if self.passes > 1:
                    raise AssertionError("style coordinates scanned more than once")
                return super().items()

        snapshot = SimpleNamespace(
            implicit_cells=lambda *_args: {},
            block=lambda *_args: {},
            note_block=lambda *_args: {},
            cell=lambda *_args: {"v": None, "s": 0, "k": "c"},
        )
        render = SimpleNamespace(
            snapshot=snapshot,
            query=SimpleNamespace(sheet=0, start=0, count=2),
            overrides={}, results={},
        )
        styles = SinglePassStyles({(0, 0): {"bold": True}, (1, 1): {"bold": True}})
        rows = WorkbookService._row_data(None, None, render, styles)
        self.assertEqual((rows["0"]["0"]["style_override"], rows["1"]["1"]["style_override"]),
                         ({"bold": True}, {"bold": True}))
        self.assertEqual(styles.passes, 1)


class BroadConditionalFormattingRowsTest(unittest.TestCase):
    def setUp(self) -> None:
        self.fixture = APIFixture(make_broad_conditional_formatting)
        self.client, self.csrf, _ = self.fixture.signup("broad-cf@example.test")
        _, copy = self.fixture.make_copy(self.client, self.csrf)
        self.copy_id = copy["id"]

    def tearDown(self) -> None:
        self.fixture.close()

    def request_rows(self, count):
        return self.client.get(
            f"/api/copies/{self.copy_id}/sheets/0/rows?start=0&count={count}",
        )

    def test_real_api_styles_normal_wide_blocks_with_one_public_engine_load(self) -> None:
        self.assertEqual(MAX_FORMULA_TARGETS, MAX_OVERRIDE_RECORDS)
        original = FormualizerEngine.load
        with patch(
            "pp.fidelity.conditional_formatting.FormualizerEngine.load",
            side_effect=original,
        ) as load:
            first = self.request_rows(50)
            full = self.request_rows(256)

        self.assertEqual((first.status_code, full.status_code), (200, 200))
        self.assertEqual(load.call_count, 2)
        for body, count in ((first.json(), 50), (full.json(), 256)):
            styled = sum(
                "style_override" in cell
                for row in body["rows"].values()
                for cell in row.values()
            )
            self.assertEqual(styled, count * 312)
            self.assertEqual(body["conditional_format_warnings"], [])
            self.assertEqual(body["rows"]["0"]["0"]["style_override"], {
                "fill": "#112233",
            })
            self.assertEqual(body["rows"][str(count - 1)]["311"]["style_override"], {
                "fill": "#112233",
            })

    def test_failed_alias_evaluation_warns_without_partial_styles(self) -> None:
        with patch(
            "pp.fidelity.conditional_formatting.FormualizerEngine.load",
            side_effect=EngineError("synthetic public-engine failure"),
        ) as load:
            response = self.request_rows(50)

        self.assertEqual(response.status_code, 200)
        body = response.json()
        self.assertEqual(load.call_count, 1)
        self.assertEqual(body["rows"], {
            "0": {"0": {"k": "c", "s": 0, "v": 1}},
        })
        self.assertEqual(
            [warning["code"] for warning in body["conditional_format_warnings"]],
            ["formula_engine"],
        )


class VolatileConditionalFormattingRowsTest(unittest.TestCase):
    def setUp(self) -> None:
        self.fixture = APIFixture(make_volatile_conditional_formatting)
        self.client, self.csrf, _ = self.fixture.signup("volatile-cf@example.test")
        _, copy = self.fixture.make_copy(self.client, self.csrf)
        self.copy_id = copy["id"]

    def tearDown(self) -> None:
        self.fixture.close()

    def test_volatile_and_named_targets_are_distinct_but_quoted_words_alias(self) -> None:
        class ContrastingEngine:
            def __init__(self):
                self.calls = []
                self.results = iter((True, False, True, True, False, True))

            def get(self, sheet, row, col):
                self.calls.append((sheet, row, col))
                return next(self.results)

        engine = ContrastingEngine()
        with patch(
            "pp.fidelity.conditional_formatting.FormualizerEngine.load",
            return_value=engine,
        ) as load:
            response = self.client.get(
                f"/api/copies/{self.copy_id}/sheets/0/rows?start=0&count=2",
            )

        self.assertEqual(response.status_code, 200)
        body = response.json()
        self.assertEqual(body["conditional_format_warnings"], [])
        self.assertEqual(body["rows"]["0"]["0"]["style_override"], {
            "fill": "#112233",
        })
        self.assertNotIn("style_override", body["rows"]["1"]["0"])
        self.assertEqual(body["rows"]["0"]["1"]["style_override"], {"bold": True})
        self.assertEqual(body["rows"]["1"]["1"]["style_override"], {"bold": True})
        self.assertEqual(body["rows"]["0"]["2"]["style_override"], {
            "fill": "#445566",
        })
        self.assertNotIn("style_override", body["rows"]["1"]["2"])
        self.assertEqual(body["rows"]["0"]["3"]["style_override"], {"bold": True})
        self.assertEqual(body["rows"]["1"]["3"]["style_override"], {"bold": True})
        self.assertEqual(load.call_count, 1)
        self.assertEqual(len(engine.calls), 6)
        self.assertTrue(all(len(call) == 3 and all(isinstance(value, int) for value in call)
                            for call in engine.calls))
