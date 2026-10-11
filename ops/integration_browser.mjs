// Run with ego-browser nodejs. Reuses the single integration TaskSpace.
const fs = await import('node:fs/promises');
const path = await import('node:path');
const assert = (await import('node:assert/strict')).default;
const { spawnSync } = await import('node:child_process');
const root = globalThis.integrationRoot;
assert.ok(path.isAbsolute(root ?? ''), 'Set integrationRoot to the synthetic worktree absolute path');
const evidence = path.join(root, 'results-private/integration');
const probe = globalThis.integrationProbe === true;
const task = await taskSpace(globalThis.integrationSpace ?? 37);
const page = task.page('p1');
const tag = Date.now();
const checks = {};
await fs.mkdir(path.join(evidence, 'browser-shots'), { recursive: true });

async function installObservers() {
  await page.cdp('Runtime.enable');
  await page.cdp('Log.enable');
  await page.cdp('Page.addScriptToEvaluateOnNewDocument', { source: `
    window.integrationCsp = [];
    window.integrationWsMessages = [];
    document.addEventListener('securitypolicyviolation', event => {
      window.integrationCsp.push({directive:event.effectiveDirective,
        blocked:event.blockedURI,source:event.sourceFile,line:event.lineNumber});
    });
    const NativeSocket = window.WebSocket;
    window.WebSocket = class extends NativeSocket {
      constructor(...args) {
        super(...args);
        this.addEventListener('message', event => {
          try { window.integrationWsMessages.push(JSON.parse(event.data)); } catch {}
        });
      }
    };
  ` });
}

async function signup() {
  await page.waitForFunction(() => document.querySelector('#auth-form') ||
    !document.querySelector('#logout-button').hidden);
  if (await page.evaluate(() => !document.querySelector('#logout-button').hidden)) {
    await page.click('#logout-button');
  }
  await page.waitForSelector('#auth-form');
  if (await page.evaluate(() => !document.querySelector('input[name=name]'))) {
    await page.click('#auth-switch');
  }
  await page.fill('input[name=name]', 'Synthetic Integration Editor');
  await page.fill('input[name=email]', `integration-${tag}@example.test`);
  await page.fill('input[name=password]', 'synthetic-integration-password');
  await page.click('#auth-form button[type=submit]');
  await page.waitForSelector('#category-create');
  checks.signup = true;
}

async function createCopy() {
  await page.click('#category-create');
  await page.fill('#modal input[name=name]', `Synthetic category ${tag}`);
  await page.click('#modal-submit');
  await page.waitForFunction(() => !document.querySelector('#copy-create').disabled);
  await page.click('#copy-create');
  const name = `Synthetic integration ${tag}`;
  await page.fill('#modal input[name=name]', name);
  await page.click('#modal-submit');
  await page.waitForFunction(name => [...document.querySelectorAll('tbody tr')]
    .some(row => row.querySelector('b')?.textContent === name), name, { timeout: 60000 });
  const id = await page.evaluate(name => [...document.querySelectorAll('tbody tr')]
    .find(row => row.querySelector('b')?.textContent === name)
    .querySelector('[data-open]').dataset.open, name);
  checks.create_copy = true;
  await page.click(`[data-open="${id}"]`);
  await page.waitForFunction(() => document.querySelector('#sheet-tabs button') &&
    document.querySelector('#grid canvas'), undefined, { timeout: 60000 });
  await page.waitForFunction(() => window.integrationWsMessages.some(message =>
    message.type === 'calc_state'), undefined, { timeout: 60000 });
  return id;
}

async function selectCell(sheet, ref) {
  await page.click(`[data-sheet="${sheet}"]`);
  await page.fill('#name-box', ref);
  await page.press('#name-box', 'Enter');
}

async function editInput(value) {
  await selectCell(0, 'B2');
  await page.waitForFunction(() => !document.querySelector('#formula-input').readOnly);
  const start = await page.evaluate(() => window.integrationWsMessages.length);
  await page.mouse.click(190, 180, { label: 'select synthetic input B2' });
  await page.keyboard.press('F2');
  await page.keyboard.press('ControlOrMeta+A');
  await page.keyboard.type(String(value));
  await page.keyboard.press('Enter');
  await page.waitForFunction(({ expected, start }) => window.integrationWsMessages.slice(start).some(message =>
    message.type === 'calc' && message.changes.some(cell =>
      cell.sheet === 2 && cell.r === 1 && cell.c === 1 && cell.v === expected)),
    { expected: value * 2, start }, { timeout: 60000 });
}

async function assertHistoryTimestamp(copyId) {
  const evidence = await page.evaluate(async id => {
    const op = window.integrationWsMessages.find(message => message.type === 'op');
    const history = await fetch(`/api/copies/${id}/history`).then(response => response.json());
    return { op, history: history.entries.find(entry => entry.op_id === op.op_id) };
  }, copyId);
  assert.equal(evidence.op.at, evidence.history.at);
  assert.ok(Number.isFinite(Date.parse(evidence.op.at)));
  checks.ws_timestamp_history = true;
  checks.native_input_ws_calc = true;
}

async function assertLocked(copyId) {
  await selectCell(2, 'B2');
  await page.waitForFunction(() => document.querySelector('#formula-input').readOnly);
  await page.mouse.click(190, 180, { label: 'focus locked formula B2' });
  await page.keyboard.press('F2');
  await page.waitForFunction(() => document.querySelector('#toast').textContent.includes('Locked'));
  const result = await page.evaluate(async id => {
    const me = await fetch('/api/me').then(response => response.json());
    const response = await fetch(`/api/copies/${id}/ops`, { method: 'POST',
      headers: { 'content-type': 'application/json', 'X-CSRF-Token': me.csrf },
      body: JSON.stringify({ op_id: 'integration-locked', action: 'edit',
        cells: [{ sheet: 2, r: 1, c: 1, v: 99, expected_version: 0 }] }) });
    return { status: response.status, body: await response.json() };
  }, copyId);
  assert.equal(result.status, 422);
  assert.equal(result.body.error, 'locked');
  checks.locked_formula_ui_and_server = true;
}

async function undo() {
  await page.click('#undo-button');
  await page.waitForFunction(() => window.integrationWsMessages.some(message =>
    message.type === 'op' && message.action === 'undo'), undefined, { timeout: 60000 });
  await selectCell(0, 'B2');
  await page.waitForFunction(() => document.querySelector('#formula-input').value === '10',
    undefined, { timeout: 60000 });
  checks.undo = true;
}

async function downloadAndRead() {
  await editInput(32);
  const pending = page.waitForEvent('download', { timeout: 30000 });
  await page.click('#export-link');
  const download = await pending;
  const filename = path.join(evidence, 'runtime/exported.xlsm');
  await download.saveAs(filename);
  const script = `import json,sys
from openpyxl import load_workbook
wb=load_workbook(sys.argv[1],keep_vba=True)
assert wb['Inputs']['B2'].value == 32
assert wb['Calc']['B2'].value == '=Inputs!B2*2'
assert wb.calculation.fullCalcOnLoad is True
print(json.dumps({'edited_value':32,'formula_preserved':True,'fullCalcOnLoad':True}))
wb.close()`;
  const result = spawnSync(path.join(root, '.venv/bin/python'), ['-c', script, filename],
    { encoding: 'utf8' });
  assert.equal(result.status, 0, result.stderr);
  checks.export_openpyxl = JSON.parse(result.stdout);
  return filename;
}

async function saveEvidence(name, extra = {}) {
  const csp = await page.evaluate(() => window.integrationCsp);
  const consoleMessages = (await page.events()).filter(event =>
    ['Runtime.consoleAPICalled', 'Runtime.exceptionThrown', 'Log.entryAdded'].includes(event.method));
  await fs.writeFile(path.join(evidence, `${name}-console.json`), JSON.stringify(consoleMessages, null, 2));
  const report = { date: '2026-10-11', synthetic: true, taskSpace: task.spaceId,
    server: 'run_v2.py', checks, csp, ...extra };
  await fs.writeFile(path.join(evidence, `${name}.json`), JSON.stringify(report, null, 2) + '\n');
  console.log(report);
  return report;
}

try {
  await installObservers();
  await page.goto('http://127.0.0.1:8765');
  console.log(await page.snapshot());
  await signup();
  const copyId = await createCopy();
  await page.screenshot({ path: path.join(evidence, 'browser-shots/workspace.png') });
  if (probe) {
    await saveEvidence('csp-before', { copyId });
    console.log(await page.snapshot());
  } else {
    await editInput(32);
    await assertHistoryTimestamp(copyId);
    await assertLocked(copyId);
    await undo();
    const download = await downloadAndRead();
    const report = await saveEvidence('browser-final', { copyId, download });
    assert.equal(report.csp.length, 0, 'CSP violations remain');
    console.log(await page.snapshot());
    if (globalThis.integrationFinish !== false) {
      await task.finish({ keep: [] });
      console.log('TaskSpace finished with keep: []');
    }
  }
} catch (error) {
  await saveEvidence(probe ? 'csp-before' : 'browser-failure', { error: String(error) });
  console.log(await page.snapshot());
  throw error;
}
