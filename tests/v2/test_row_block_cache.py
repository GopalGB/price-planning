"""A rendered row block is reused until an input the body depends on changes."""
import json
import unittest
from unittest import mock

from pp.calc_service import CalculationPublication, CalculationService
from pp.pivot_state import write_state
from pp.row_cache import RowBlockCache
from tests.v2.api_support import APIFixture


class RowBlockCacheBudgetTests(unittest.TestCase):
    """The cache is bounded by bytes and by entry count, and drops the least recently used entry."""

    BLOCK = {"rows": {"1": {"1": {"v": "0123456789"}}}}
    SIZE = len(json.dumps(BLOCK, separators=(",", ":")))

    def test_the_byte_budget_drops_the_oldest_entry(self):
        cache = RowBlockCache(max_bytes=self.SIZE * 2 + 1, max_entries=10)
        for key in ("a", "b"):
            cache.put(key, self.BLOCK)
        self.assertEqual(cache.bytes_held, self.SIZE * 2)
        cache.put("c", self.BLOCK)
        self.assertIsNone(cache.get("a"))
        self.assertEqual(cache.get("c"), self.BLOCK)
        self.assertLessEqual(cache.bytes_held, self.SIZE * 2 + 1)

    def test_the_entry_budget_drops_the_oldest_entry(self):
        cache = RowBlockCache(max_bytes=1024, max_entries=2)
        for key in ("a", "b", "c"):
            cache.put(key, self.BLOCK)
        self.assertIsNone(cache.get("a"))
        self.assertEqual(cache.entries, 2)

    def test_a_read_refreshes_an_entry_it_keeps(self):
        cache = RowBlockCache(max_bytes=1024, max_entries=2)
        for key in ("a", "b"):
            cache.put(key, self.BLOCK)
        cache.get("a")
        cache.put("c", self.BLOCK)
        self.assertIsNone(cache.get("b"))
        self.assertEqual(cache.get("a"), self.BLOCK)

    def test_replacing_an_entry_does_not_double_count_its_bytes(self):
        cache = RowBlockCache(max_bytes=1024, max_entries=10)
        cache.put("a", self.BLOCK)
        cache.put("a", self.BLOCK)
        self.assertEqual((cache.entries, cache.bytes_held), (1, self.SIZE))

    def test_a_body_larger_than_the_whole_budget_is_not_stored(self):
        cache = RowBlockCache(max_bytes=self.SIZE - 1, max_entries=10)
        cache.put("a", self.BLOCK)
        self.assertIsNone(cache.get("a"))
        self.assertEqual(cache.bytes_held, 0)

    def test_clearing_releases_every_entry(self):
        cache = RowBlockCache(max_bytes=1024, max_entries=10)
        cache.put("a", self.BLOCK)
        cache.clear()
        self.assertEqual((cache.entries, cache.bytes_held), (0, 0))


class RowBlockCacheContractTests(unittest.TestCase):
    """Every input the rendered body depends on has to be part of the cache key."""

    BLOCK = "start=0&count=2"

    def setUp(self):
        self.fixture = APIFixture()
        self.client, self.csrf, self.user = self.fixture.signup("cache@example.test", "Cache Owner")
        _, self.copy = self.fixture.make_copy(self.client, self.csrf)
        self.services = self.fixture.app.state.services
        self.original = self.services.workbooks.row_cache
        self.renders = self._count_renders()
        self.addCleanup(setattr, self.services.workbooks, "row_cache", self.original)

    def tearDown(self):
        self.fixture.close()

    def _count_renders(self):
        from pp import workbook_service

        calls = []
        original = workbook_service.conditional_style_overrides

        def counted(*args, **kwargs):
            calls.append(1)
            return original(*args, **kwargs)

        patcher = mock.patch.object(workbook_service, "conditional_style_overrides", counted)
        patcher.start()
        self.addCleanup(patcher.stop)
        return calls

    def rows(self, query=BLOCK, sheet=0, client=None):
        return (client or self.client).get(
            f"/api/copies/{self.copy['id']}/sheets/{sheet}/rows?{query}")

    def test_an_unchanged_copy_is_served_from_the_cache(self):
        first = self.rows()
        second = self.rows()
        self.assertEqual(first.status_code, 200)
        self.assertEqual(len(self.renders), 1)
        self.assertEqual(first.json(), second.json())

    def test_another_range_of_the_same_sheet_is_its_own_block(self):
        self.rows()
        self.rows("start=0&count=3")
        self.assertEqual(len(self.renders), 2)

    def test_an_edit_rerenders_the_block(self):
        self.rows()
        self.services.operations.apply(self.copy["id"], self.user["id"], {
            "op_id": "cache-edit", "action": "edit",
            "cells": [{"sheet": 0, "r": 1, "c": 1, "v": 4242, "expected_version": 0}],
        })
        body = self.rows().json()
        self.assertEqual(len(self.renders), 2)
        self.assertEqual(body["rows"]["1"]["1"]["v"], 4242)
        self.assertEqual(body["input_revision"], 1)
        self.rows()
        self.assertEqual(len(self.renders), 2)

    def test_a_published_calculation_rerenders_the_block(self):
        self.rows()
        self.services.operations.apply(self.copy["id"], self.user["id"], {
            "op_id": "cache-pending", "action": "edit",
            "cells": [{"sheet": 0, "r": 1, "c": 1, "v": 12, "expected_version": 0}],
        })
        self.rows()
        CalculationService(self.fixture.database).complete(CalculationPublication(
            self.copy["id"], 1, [(2, 1, 1, 222)],
        ))
        body = self.rows().json()
        calculated = self.rows(sheet=2).json()
        self.assertEqual(len(self.renders), 4)
        self.assertEqual((body["input_revision"], body["calculated_revision"]), (1, 1))
        self.assertEqual(body["rows"]["1"]["1"]["v"], 12)
        self.assertEqual(calculated["rows"]["1"]["1"]["v"], 222)
        self.assertEqual(len(self.renders), 4)

    def test_a_stored_slicer_selection_rerenders_the_block(self):
        self.rows()
        with self.fixture.database.transaction() as connection:
            write_state(connection, self.copy["id"],
                        {"slicers": {"Slicer_Region": ["North"]}, "refreshed": False}, 1)
        self.rows()
        self.assertEqual(len(self.renders), 2)
        self.rows()
        self.assertEqual(len(self.renders), 2)

    def test_access_is_checked_before_the_cache_is_read(self):
        self.rows()
        outsider, _, _ = self.fixture.signup("outsider@example.test", "Outsider")
        response = self.rows(client=outsider)
        self.assertEqual((response.status_code, response.json()["error"]), (404, "not_found"))
        self.assertEqual(len(self.renders), 1)

    def test_another_member_of_the_same_copy_reads_the_same_block(self):
        first = self.rows().json()
        member, _, _ = self.fixture.signup("member@example.test", "Copy Member")
        shared = self.client.put(f"/api/copies/{self.copy['id']}/shares", {
            "email": "member@example.test", "role": "viewer",
        }, {"x-csrf-token": self.csrf})
        self.assertEqual(shared.status_code, 200, shared.body)
        response = self.rows(client=member)
        self.assertEqual(response.status_code, 200)
        self.assertEqual(response.json(), first)
        self.assertEqual(len(self.renders), 1)


if __name__ == "__main__":
    unittest.main()