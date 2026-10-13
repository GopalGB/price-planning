import test from "node:test";
import assert from "node:assert/strict";
import { drawingsView, linkTarget } from "../src/drawings.mjs";

const sheet = { index: 0, name: "Main", col_widths: { 0: 100 }, row_heights: { 0: 20 } };
const meta = {
  sheets: [sheet, { index: 3, name: "Plan 'A'" }],
  drawings: [
    {
      id: "d0", sheet: 0, kind: "picture",
      from: { row: 1, col: 1, dx: 2, dy: 3 }, size: { w: 40, h: 30 },
      src: "data:image/png;base64,AAAA",
    },
    {
      id: "d1", sheet: 0, kind: "shape", geom: "roundRect", adjust: 0.25,
      from: { row: 0, col: 0, dx: 0, dy: 0 }, to: { row: 1, col: 1, dx: 0, dy: 0 },
      fill: "#4472C4", line: { color: "#2F528F", width: 1 }, shadow: true,
      link: { sheet: "Plan 'A'", cell: "B7" },
      text: {
        anchor: "ctr", wrap: true, inset: [4, 2, 4, 2],
        paragraphs: [{ align: "ctr", runs: [{ t: "Go <now>", b: true, i: false, sz: 16, color: "#FFFFFF", font: "Calibri" }] }],
      },
    },
    { id: "d2", sheet: 0, kind: "picture", from: { row: 0, col: 0, dx: 0, dy: 0 }, size: { w: 1, h: 1 }, src: "javascript:alert(1)" },
    { id: "d3", sheet: 0, kind: "shape", geom: "downArrow", from: { row: 0, col: 0, dx: 0, dy: 0 }, size: { w: 20, h: 40 }, fill: "red;x", line: null },
    { id: "d4", sheet: 3, kind: "picture", from: { row: 0, col: 0, dx: 0, dy: 0 }, size: { w: 1, h: 1 }, src: "data:image/png;base64,BB" },
  ],
};

test("pictures and shapes sit where Excel anchors them, minus the scroll", () => {
  const html = drawingsView(meta, 0, { scroll: { sheetViewStartRow: 0, sheetViewStartColumn: 0, offsetX: 10, offsetY: 0 } });
  // column 0 is 100px wide, row 0 is 20px high: picture at (100+2-10, 20+3)
  assert.match(html, /<img class="sheet-picture" alt="" src="data:image\/png;base64,AAAA" style="left:92px;top:23px;width:40px;height:30px">/);
  // shape spans A1 to B2's top-left corner: 100 x 20, radius 25% of the shorter side
  assert.match(html, /left:-10px;top:0px;width:100px;height:20px;background:#4472C4;border:1px solid #2F528F;border-radius:5px;box-shadow/);
  assert.match(html, /data-link-sheet="Plan &#39;A&#39;" data-link-cell="B7"/);
  assert.match(html, /justify-content:center;padding:2px 4px 2px 4px/);
  assert.match(html, /text-align:center"><span style="font-family:'Calibri';font-size:16px;color:#FFFFFF;font-weight:700;font-style:normal">Go &lt;now&gt;<\/span>/);
  assert.doesNotMatch(html, /javascript/);
  assert.match(html, /<polygon points="5,0 15,0 15,30 20,30 10,40 0,30 5,30" fill="transparent" stroke="none"/);
  assert.doesNotMatch(html, /BB/);
});

test("a shape link resolves to the sheet index by name", () => {
  const element = { dataset: { linkSheet: "Plan 'A'", linkCell: "B7" } };
  assert.deepEqual(linkTarget(meta, element), { sheet: 3, cell: "B7" });
  assert.equal(linkTarget(meta, { dataset: { linkSheet: "Gone", linkCell: "A1" } }), null);
  assert.equal(linkTarget(meta, null), null);
});

test("with grid geometry, overlays use Univer cell positions, headers and zoom", async () => {
  const { macroButtonsView } = await import("../src/macros.mjs");
  const geometry = {
    cellRect: (r, c) => ({ left: 40 + c * 64, right: 104 + c * 64, top: 20 + r * 20, bottom: 40 + r * 20 }),
    zoom: 2,
    headers: { left: 40, top: 20 },
    freeze: { rows: 0, cols: 0 },
    scroll: { viewportScrollX: 0, viewportScrollY: 0 },
  };
  const item = { id: "d9", sheet: 0, kind: "picture", from: { row: 1, col: 1, dx: 5, dy: 0 }, size: { w: 10, h: 10 }, src: "data:image/png;base64,CC" };
  const html = drawingsView({ sheets: [sheet], drawings: [item] }, 0, { geometry });
  assert.match(html, /style="left:218px;top:80px;width:10px;height:10px;transform:scale\(2\);transform-origin:0 0"/);
  const button = {
    id: "b0", sheet: 0, label: "Run", from: { row: 0, col: 0, dx: 0, dy: 0 }, to: { row: 1, col: 1, dx: 0, dy: 0 },
    look: { geom: "roundRect", adjust: 0.5, fill: "#ED7D31", line: null, shadow: false,
      text: { anchor: "ctr", wrap: true, inset: [0, 0, 0, 0], paragraphs: [{ align: "ctr", runs: [{ t: "Run", b: false, i: false, sz: 12, color: "#FFFFFF" }] }] } },
  };
  const buttons = macroButtonsView({ sheets: [sheet], buttons: [button] }, 0, { geometry, copyRole: "editor" });
  // A1 to B2: 64 x 20 unzoomed, corners half of the shorter side
  assert.match(buttons, /class="macro-button macro-shape"[^>]*style="left:80px;top:40px;width:64px;height:20px;transform:scale\(2\);transform-origin:0 0;background:#ED7D31;border-radius:10px"/);
  assert.match(buttons, /<div class="sheet-shape-text"[^>]*><p style="text-align:center"><span[^>]*>Run<\/span>/);
});
