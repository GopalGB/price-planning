import assert from "node:assert/strict";
import test from "node:test";

// Office users open the app over plain http on the LAN, where browsers hide crypto.randomUUID.
test("operation ids work without crypto.randomUUID (insecure LAN context)", async () => {
  const original = crypto.randomUUID;
  Object.defineProperty(crypto, "randomUUID", {
    value: undefined,
    configurable: true,
  });
  try {
    const { newOpId, buildOp, invertOp } = await import(
      `../src/model.mjs?ids=${Date.now()}`
    );
    const ids = new Set(Array.from({ length: 200 }, () => newOpId()));
    assert.equal(ids.size, 200);
    for (const id of ids)
      assert.match(
        id,
        /^[0-9a-f]{8}-[0-9a-f]{4}-4[0-9a-f]{3}-[89ab][0-9a-f]{3}-[0-9a-f]{12}$/,
      );
    assert.equal(typeof buildOp, "function");
    assert.equal(typeof invertOp, "function");
  } finally {
    Object.defineProperty(crypto, "randomUUID", {
      value: original,
      configurable: true,
    });
  }
});
