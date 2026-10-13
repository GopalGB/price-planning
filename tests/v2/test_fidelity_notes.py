import json
import tempfile
import unittest
from pathlib import Path

from ops.fixtures import NOTE_TEXT, make_notes
from pp.template.importer import import_template
from pp.workbook_service import TemplateSnapshot
from tests.v2.api_support import APIFixture


class NoteImporterTests(unittest.TestCase):
    def test_imports_two_comment_parts_as_bounded_note_blocks(self):
        with tempfile.TemporaryDirectory() as directory:
            source = Path(directory) / "notes.xlsx"
            make_notes(source)
            path = import_template(source, Path(directory) / "data")
            snapshot = TemplateSnapshot(path)

            self.assertTrue(snapshot.sheets[0]["has_notes"])
            self.assertTrue(snapshot.sheets[1]["has_notes"])
            self.assertGreaterEqual(snapshot.sheets[0]["rows"], 300)
            self.assertGreaterEqual(snapshot.sheets[1]["cols"], 4)
            self.assertEqual(snapshot.note_block(0, 0)["1"]["1"], {
                "author": "A <Admin>", "text": NOTE_TEXT,
            })
            self.assertEqual(snapshot.note_block(0, 1)["43"]["1"]["text"], "Far block")
            self.assertEqual(snapshot.note_block(1, 0)["3"]["3"], {
                "author": "Reviewer", "text": "Second line continues",
            })
            meta = json.loads((path / "meta.json").read_text())
            self.assertNotIn(NOTE_TEXT, json.dumps(meta))


class NoteAPIContractTests(unittest.TestCase):
    def setUp(self):
        self.fixture = APIFixture(make_notes)
        self.owner, self.csrf, _ = self.fixture.signup("note-owner@example.test")
        _, self.copy = self.fixture.make_copy(self.owner, self.csrf)

    def tearDown(self):
        self.fixture.close()

    def test_rows_stream_blank_cell_notes_from_the_requested_blocks(self):
        copy_id = self.copy["id"]
        workbook = self.owner.get(f"/api/copies/{copy_id}/workbook").json()
        self.assertNotIn(NOTE_TEXT, json.dumps(workbook))

        first = self.owner.get(
            f"/api/copies/{copy_id}/sheets/0/rows?start=1&count=1"
        ).json()
        far = self.owner.get(
            f"/api/copies/{copy_id}/sheets/0/rows?start=299&count=1"
        ).json()
        second = self.owner.get(
            f"/api/copies/{copy_id}/sheets/1/rows?start=3&count=1"
        ).json()

        self.assertIsNone(first["rows"]["1"]["1"]["v"])
        self.assertEqual(first["rows"]["1"]["1"]["note"], {
            "author": "A <Admin>", "text": NOTE_TEXT,
        })
        self.assertEqual(far["rows"]["299"]["1"]["note"]["text"], "Far block")
        self.assertEqual(second["rows"]["3"]["3"]["note"]["author"], "Reviewer")

    def test_rows_do_not_read_an_unrequested_note_block(self):
        path = self.fixture.snapshot_path / "notes" / "0" / "1.json"
        path.write_text("not json", encoding="utf-8")

        response = self.owner.get(
            f"/api/copies/{self.copy['id']}/sheets/0/rows?start=1&count=1"
        )

        self.assertEqual(response.status_code, 200)
        self.assertEqual(response.json()["rows"]["1"]["1"]["note"]["text"], NOTE_TEXT)


if __name__ == "__main__":
    unittest.main()
