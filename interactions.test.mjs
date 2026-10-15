import { test } from 'node:test';
import assert from 'node:assert/strict';
import { readFileSync } from 'node:fs';

// A small DOM harness checks application event handling and download payloads.
// It does not render CSS or substitute for browser visual QA.
class MockNode {
  constructor(tag, attributes = {}) {
    this.tagName = tag; this.attributes = attributes; this.value = attributes.value || '';
    this.checked = 'checked' in attributes; this.hidden = 'hidden' in attributes;
    this.children = []; this.listeners = {}; this.style = {}; this.classes = new Set();
    this.classList = { toggle: (name, on) => on ? this.classes.add(name) : this.classes.delete(name) };
  }
  set textContent(v) { this.text = String(v); this.children = []; }
  get textContent() { return (this.text || '') + this.children.map(x => x.textContent).join(''); }
  append(...nodes) { this.children.push(...nodes); }
  replaceChildren(...nodes) { this.text = ''; this.children = nodes; }
  setAttribute(k, v) { this.attributes[k] = v; }
  removeAttribute(k) { delete this.attributes[k]; }
  addEventListener(event, callback) { (this.listeners[event] ||= []).push(callback); }
  dispatch(event, data = {}) { for (const cb of this.listeners[event] || []) cb({ preventDefault() {}, ...data }); }
  focus() {} scrollIntoView() {} remove() {} showModal() { this.open = true; } close() { this.open = false; }
  click() { this.dispatch('click'); if (this.tagName === 'a') downloads.push({ filename: this.download, blob: blobs.get(this.href) }); }
}
function attributes(text) {
  const a = {};
  for (const m of text.matchAll(/([\w-]+)(?:="([^"]*)")?/g)) a[m[1]] = m[2] ?? '';
  return a;
}
const html = readFileSync(new URL('../dist/index.html', import.meta.url), 'utf8');
const nodes = new Map(), controls = [], downloads = [], blobs = new Map();
for (const m of html.matchAll(/<([a-z][a-z0-9-]*)\b([^>]*)>/g)) {
  const a = attributes(m[2]), node = new MockNode(m[1], a);
  if (a.id) nodes.set(a.id, node);
  if (['input', 'select', 'textarea'].includes(m[1])) controls.push(node);
}
for (const m of html.matchAll(/<select\b([^>]*)>([\s\S]*?)<\/select>/g)) {
  const node = nodes.get(attributes(m[1]).id);
  const options = [...m[2].matchAll(/<option\b([^>]*)>/g)].map(x => attributes(x[1]));
  node.value = (options.find(x => 'selected' in x) || options[0]).value;
}
function query(selector) {
  const name = selector.match(/name="([^"]+)"/)?.[1];
  return controls.filter(n => n.attributes.name === name && (!selector.endsWith(':checked') || n.checked));
}
nodes.get('config-form').querySelectorAll = query;
nodes.get('config-form').querySelector = s => query(s)[0];
const stepNav = new MockNode('nav');
const body = new MockNode('body');

test('application interactions validate review, update dependencies, reset and produce real downloads', async t => {
  const original = { document: globalThis.document, localStorage: globalThis.localStorage, setTimeout: globalThis.setTimeout, clearTimeout: globalThis.clearTimeout, createObjectURL: URL.createObjectURL, revokeObjectURL: URL.revokeObjectURL };
  globalThis.document = { getElementById: id => nodes.get(id), createElement: tag => new MockNode(tag), querySelector: () => stepNav, body };
  const saved = new Map(); globalThis.localStorage = { getItem: k => saved.get(k) || null, setItem: (k, v) => saved.set(k, v) };
  globalThis.setTimeout = () => 1; globalThis.clearTimeout = () => {};
  URL.createObjectURL = blob => { const key = 'blob:mock-' + blobs.size; blobs.set(key, blob); return key; };
  URL.revokeObjectURL = () => {};
  const get = id => nodes.get(id), input = () => get('config-form').dispatch('input');
  const radio = (name, value) => { for (const n of controls.filter(n => n.attributes.name === name)) n.checked = n.value === value; input(); };
  try {
    await import('../dist/app.js?interaction-test');
    await t.test('default review and JSON download contain both formats and no fabricated counts', async () => {
      get('review-tab').click(); assert.equal(get('review-panel').hidden, false); assert.equal(get('configure-panel').hidden, true);
      assert.match(get('review-summary').textContent, /Entire assembly/);
      get('download-config').click(); const download = downloads.at(-1); assert.equal(download.filename, 'axle_exterior_configuration.json');
      const p = JSON.parse(await download.blob.text()); assert.equal(p.export.files.length, 2); assert.equal(p.expectedResultReport.geometry.bodyCount, null);
      assert.match(get('download-status').textContent, /No CAD geometry/);
    });
    await t.test('scope validation blocks review until explicit occurrence selection exists', () => {
      get('edit').click(); get('assemblyScope').value = 'components'; input(); get('side-review').click();
      assert.equal(get('review-panel').hidden, true); assert.equal(get('form-errors').hidden, false);
      get('scopeSelection').value = 'ROOT/HOUSING_1\nROOT/COVER_1'; input(); get('side-review').click();
      assert.equal(get('review-panel').hidden, false); assert.equal(get('form-errors').hidden, true);
      assert.match(get('review-summary').textContent, /ROOT\/COVER_1/);
    });
    await t.test('solid mode enforces required switches and existing mode skips geometry processing', async () => {
      get('edit').click(); radio('geometryMode', 'solid');
      assert.equal(get('sewSurfaces').checked, true); assert.equal(get('attemptSolid').checked, true); assert.equal(get('sewSurfaces').disabled, true);
      get('existingExterior').value = 'existing'; input(); assert.equal(get('geometry-fields').disabled, true);
      get('review-tab').click(); get('download-config').click(); const p = JSON.parse(await downloads.at(-1).blob.text());
      assert.equal(p.geometry.mode, 'existing'); assert.equal(p.geometry.attemptSolidWhenClosed, false);
    });
    await t.test('custom Parasolid binary plan and readable report reflect current settings', async () => {
      get('edit').click(); radio('outputFormat', 'parasolid'); get('parasolidVersion').value = 'custom'; input();
      assert.equal(get('target-version-field').hidden, false); get('review-tab').click(); assert.equal(get('form-errors').hidden, false);
      get('targetVersion').value = '32.0'; get('parasolidType').value = 'binary'; get('outputName').value = 'carrier_v2'; input();
      get('review-tab').click(); get('download-config').click(); const p = JSON.parse(await downloads.at(-1).blob.text());
      assert.equal(p.export.files[0].filename, 'carrier_v2.x_b'); assert.equal(p.export.parasolid.version, '32.0');
      assert.match(get('summary-files').textContent, /X_B/);
      get('download-report').click(); const r = downloads.at(-1); assert.equal(r.filename, 'carrier_v2_expected_report.txt');
      assert.match(await r.blob.text(), /carrier_v2\.x_b/); assert.match(await r.blob.text(), /NOT EXECUTED/);
    });
    await t.test('reset restores defaults and help dialog opens and closes', () => {
      get('edit').click(); get('reset').click(); assert.equal(get('assemblyScope').value, 'entire'); assert.equal(get('outputName').value, 'axle_exterior');
      assert.equal(get('geometry-fields').disabled, false); assert.equal(get('attemptSolid').checked, false);
      get('help').click(); assert.equal(get('help-dialog').open, true); get('close-help').click(); assert.equal(get('help-dialog').open, false);
      assert.ok(saved.size > 0);
    });
  } finally {
    globalThis.document = original.document; globalThis.localStorage = original.localStorage;
    globalThis.setTimeout = original.setTimeout; globalThis.clearTimeout = original.clearTimeout;
    URL.createObjectURL = original.createObjectURL; URL.revokeObjectURL = original.revokeObjectURL;
  }
});
