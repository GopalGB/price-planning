from types import SimpleNamespace

from ops.fixtures import make_semantic
from tests.v2.asgi_client import ASGIClient
from tests.v2.support import BackendFixture


class APIFixture(BackendFixture):
    def __init__(self, builder=make_semantic):
        super().__init__(builder)
        from pp.engine.pool import EnginePool
        from pp.server import ServiceContext, create_app
        self.config = SimpleNamespace(
            demo_mode=True, allowed_origins=("http://testserver",), cookie_secure=False,
            replay_window=100, max_workers=1, template_path=self.source,
            data_dir=self.root / "data", bind="127.0.0.1", port=8000,
        )
        self.pool = EnginePool(max_workers=1)
        self.app = create_app(ServiceContext(self.database, self.snapshot, self.config, self.pool))

    def client(self):
        return ASGIClient(self.app)

    def signup(self, email, name="Test User", password="correct-horse-battery"):
        client = self.client()
        response = client.post("/api/auth/signup", {"email": email, "name": name, "password": password})
        if response.status_code != 201:
            raise AssertionError(response.body)
        me = client.get("/api/me")
        return client, me.json()["csrf"], me.json()["user"]

    def make_copy(self, client, csrf, name="Plan", visibility="private", default_role="viewer"):
        category = client.post("/api/categories", {"name": f"{name} Category"},
                               {"x-csrf-token": csrf}).json()["category"]
        copy = client.post("/api/copies", {
            "category_id": category["id"], "name": name,
            "visibility": visibility, "default_role": default_role,
        }, {"x-csrf-token": csrf}).json()["copy"]
        return category, copy

    def close(self):
        self.pool.close()
        super().close()
