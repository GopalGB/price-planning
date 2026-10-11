// Univer gate, run in Ego Lite: `npm run build && npx vite preview --port 4310 &` then `ego-browser nodejs < gate-ego.mjs`
const { writeFileSync } = await import("node:fs");
const OUT = "/Users/gopalmacbook/Desktop/Max HQ/price-planning/.planning/excel-grid/results/univer-gate-2026-10-10.json";
const task = await taskSpace("univer gate");
const page = task.page("p1");
const out = { univer: "1.0.4", browser: "Ego Lite", heap_limit_mb: 1500, sheets: {} };
for (const [sheet, mode] of [["grid", "windowed"], ["tall", "windowed"], ["grid", "full"]]) {
  await page.goto(`http://127.0.0.1:4310/?sheet=${sheet}&mode=${mode}`);
  await page.waitForFunction(() => !!(window.gate && window.gate.statistics), undefined, { timeout: 120000 });
  await page.waitForTimeout(1000);
  await page.cdp("HeapProfiler.collectGarbage");
  const r = { heap_mb: Math.round((await page.cdp("Runtime.getHeapUsage")).usedSize / 1e6) };
  await page.evaluate(() => { window.gate.select(8, 1); });
  const before8 = await page.evaluate(() => String(window.gate.value(8, 1)));
  const before9 = await page.evaluate(() => String(window.gate.value(9, 1)));
  await page.keyboard.type("=SUM(1,2)"); await page.keyboard.press("Enter"); await page.waitForTimeout(600);
  r.blocked = await page.evaluate(() => window.gate.blockedEdits());
  r.blocked_cell_unchanged = await page.evaluate((b) => String(window.gate.value(8, 1)) === b, before8);
  await page.evaluate(() => { window.gate.select(9, 1); });
  await page.keyboard.type("77"); await page.keyboard.press("Enter"); await page.waitForTimeout(600);
  r.number_saved = await page.evaluate(() => String(window.gate.value(9, 1)) === "77");
  await page.evaluate(async () => { await window.gate.undo(); }); await page.waitForTimeout(400);
  r.undo_restored = await page.evaluate((b) => String(window.gate.value(9, 1)) === b, before9);
  r.formulas = await page.evaluate(() => window.gate.formulaCellCount());
  r.pass = r.heap_mb < 1500 && r.blocked === 1 && r.blocked_cell_unchanged && r.number_saved && r.undo_restored && r.formulas === 0;
  out.sheets[`${sheet}_${mode}`] = r;
}
out.pass = Object.values(out.sheets).every((s) => s.pass);
out.not_measured = ["scroll fps (Ego background tab throttles rAF; measure in P3 with a visible window)", "paste interception (BeforeClipboardPaste) and fill-handle interception: P3"];
writeFileSync(OUT, JSON.stringify(out, null, 2) + "\n");
console.log(JSON.stringify(out, null, 1));
await task.finish({ keep: [] });
