import unittest

from tests.v2.api_support import APIFixture


class CategoryAPIContractTests(unittest.TestCase):
    def setUp(self):
        self.fixture = APIFixture()
        self.manager, self.manager_csrf, _ = self.fixture.signup("manager@example.test", "Manager")
        self.member, self.member_csrf, _ = self.fixture.signup("member@example.test", "Member")

    def tearDown(self):
        self.fixture.close()

    def test_manager_can_add_member_and_member_cannot_manage(self):
        """Category membership changes without manager authorization must fail."""
        created = self.manager.post("/api/categories", {"name": "Category A"},
                                    {"x-csrf-token": self.manager_csrf})
        self.assertEqual(created.status_code, 201)
        category_id = created.json()["category"]["id"]
        updated = self.manager.put(f"/api/categories/{category_id}/members",
                                   {"email": "member@example.test", "role": "member"},
                                   {"x-csrf-token": self.manager_csrf})
        self.assertEqual(updated.status_code, 200)
        listed = self.member.get("/api/categories").json()["categories"]
        self.assertEqual(listed[0]["role"], "member")
        self.assertEqual({member["email"] for member in listed[0]["members"]},
                         {"manager@example.test", "member@example.test"})
        forbidden = self.member.put(f"/api/categories/{category_id}/members",
                                    {"email": "manager@example.test", "role": None},
                                    {"x-csrf-token": self.member_csrf})
        self.assertEqual((forbidden.status_code, forbidden.json()["error"]), (403, "forbidden"))


if __name__ == "__main__":
    unittest.main()

