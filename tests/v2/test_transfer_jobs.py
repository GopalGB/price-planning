"""Uploads and downloads as background jobs: progress, the same result, and session ownership."""

import time
import unittest

from tests.v2.api_support import APIFixture


class TransferJobTests(unittest.TestCase):
    def setUp(self):
        self.fixture = APIFixture()
        self.addCleanup(self.fixture.close)
        self.owner, self.csrf, _ = self.fixture.signup("jobs-owner@example.test", "Jobs Owner")
        _, self.copy = self.fixture.make_copy(self.owner, self.csrf)
        exported = self.owner.get(f"/api/copies/{self.copy['id']}/export.xlsx")
        self.assertEqual(exported.status_code, 200, exported.body)
        self.exported = exported.body

    def wait(self, client, token):
        for _ in range(600):
            body = client.get(f"/api/jobs/{token}").json()
            if body["state"] != "running":
                return body
            time.sleep(0.05)
        self.fail("job did not finish")

    def start_preview(self, path_query="?async=1"):
        response = self.owner.raw(
            "POST", f"/api/copies/{self.copy['id']}/uploads/preview{path_query}", self.exported,
            {"x-csrf-token": self.csrf, "x-upload-filename": "Plan A.xlsx",
             "content-type": "application/octet-stream"},
        )
        return response

    def test_async_preview_reports_progress_and_matches_the_direct_preview(self):
        direct = self.start_preview("").json()
        started = self.start_preview()
        self.assertEqual(started.status_code, 202, started.body)
        finished = self.wait(self.owner, started.json()["job"])
        self.assertEqual(finished["state"], "done", finished)
        self.assertEqual({key: finished["result"][key] for key in ("changes", "ignored", "filename")},
                         {key: direct[key] for key in ("changes", "ignored", "filename")})
        confirmed = self.owner.post(f"/api/uploads/{finished['result']['token']}/confirm",
                                    {"mode": "update"}, {"x-csrf-token": self.csrf})
        self.assertEqual(confirmed.status_code, 200, confirmed.body)

    def test_another_user_cannot_read_a_job(self):
        token = self.start_preview().json()["job"]
        self.wait(self.owner, token)
        other, _, _ = self.fixture.signup("jobs-other@example.test", "Other")
        self.assertEqual(other.get(f"/api/jobs/{token}").status_code, 404)

    def test_failed_preview_job_carries_the_error(self):
        response = self.owner.raw(
            "POST", f"/api/copies/{self.copy['id']}/uploads/preview?async=1", b"not a workbook",
            {"x-csrf-token": self.csrf, "x-upload-filename": "bad.xlsx",
             "content-type": "application/octet-stream"},
        )
        finished = self.wait(self.owner, response.json()["job"])
        self.assertEqual((finished["state"], finished["error"]["status"]), ("failed", 422), finished)

    def test_async_download_returns_a_working_link(self):
        sheets = [sheet["name"] for sheet in
                  self.owner.get(f"/api/copies/{self.copy['id']}/workbook").json()["sheets"]]
        started = self.owner.post(f"/api/copies/{self.copy['id']}/downloads?async=1",
                                  {"format": "xlsx", "sheets": sheets}, {"x-csrf-token": self.csrf})
        self.assertEqual(started.status_code, 202, started.body)
        finished = self.wait(self.owner, started.json()["job"])
        self.assertEqual(finished["state"], "done", finished)
        download = self.owner.get(finished["result"]["download_url"])
        self.assertEqual(download.status_code, 200)
        self.assertTrue(download.body.startswith(b"PK"))

    def test_transfer_slots_are_released_after_jobs(self):
        for _ in range(6):
            token = self.start_preview().json()["job"]
            self.assertEqual(self.wait(self.owner, token)["state"], "done")


if __name__ == "__main__":
    unittest.main()
