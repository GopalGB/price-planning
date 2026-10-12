import unittest

from tests.v2.api_support import APIFixture
from tests.v2.test_macro_buttons import write_button_fixture


class MacroEndpointContractTests(unittest.TestCase):
    def setUp(self):
        self.fixture = APIFixture(builder=write_button_fixture)
        self.client, self.csrf, _ = self.fixture.signup("macros@example.test", "Macro Owner")
        _, self.copy = self.fixture.make_copy(self.client, self.csrf)
        self.workbook = self.client.get(f"/api/copies/{self.copy['id']}/workbook").json()
        self.copy_id = self.copy["id"]

    def tearDown(self):
        self.fixture.close()

    def test_workbook_payload_lists_buttons_with_position_and_macro_only(self):
        buttons = self.workbook["buttons"]
        self.assertEqual([button["id"] for button in buttons], ["b0", "b1", "b2"])
        self.assertEqual(set(buttons[0]), {"id", "sheet", "from", "to", "label", "macro"})
        self.assertNotIn("buttons", self.workbook["sheets"][0])

    def test_owner_pressing_known_button_gets_macros_unavailable_until_office_run_exists(self):
        response = self.client.post(f"/api/copies/{self.copy_id}/macros", {"button": "b2", "op_id": "press-1"},
                                    {"x-csrf-token": self.csrf})
        self.assertEqual(response.status_code, 409)
        self.assertEqual(response.json()["error"], "macros_unavailable")
        self.assertEqual(response.json()["message"],
                         "Macros run on the office server, where the app uses Microsoft Excel.")

    def test_unknown_button_is_404(self):
        response = self.client.post(f"/api/copies/{self.copy_id}/macros", {"button": "b9", "op_id": "press-2"},
                                    {"x-csrf-token": self.csrf})
        self.assertEqual((response.status_code, response.json()["error"]), (404, "unknown_button"))

    def test_viewer_is_forbidden_even_for_known_button(self):
        viewer, viewer_csrf, _ = self.fixture.signup("macro-viewer@example.test", "Viewer")
        self.client.put(f"/api/copies/{self.copy_id}/shares", {"email": "macro-viewer@example.test", "role": "viewer"},
                        {"x-csrf-token": self.csrf})
        response = viewer.post(f"/api/copies/{self.copy_id}/macros", {"button": "b2", "op_id": "press-3"},
                               {"x-csrf-token": viewer_csrf})
        self.assertEqual((response.status_code, response.json()["error"]), (403, "forbidden"))

    def test_missing_csrf_and_anonymous_requests_are_rejected(self):
        anonymous = self.fixture.client()
        self.assertEqual(anonymous.post(f"/api/copies/{self.copy_id}/macros", {"button": "b2", "op_id": "x"},
                                        {"x-csrf-token": self.csrf}).status_code, 401)
        self.assertEqual(self.client.post(f"/api/copies/{self.copy_id}/macros", {"button": "b2", "op_id": "x"}
                                          ).json()["error"], "csrf")


if __name__ == "__main__":
    unittest.main()
