import { freshDefaults, normalize, validate, outputs, warnings, summary, packageConfig, LABELS } from './config.js';

const $ = id => document.getElementById(id);
const form = $('config-form');
const STORAGE_KEY = 'nx2306-exterior-export-v1';
const checkboxes = ['visibleOnly', 'excludeInternal', 'closeOpenings', 'sewSurfaces', 'attemptSolid', 'previewIsolation'];
const strings = ['assemblyName', 'assemblyScope', 'scopeSelection', 'customExclusions', 'quality', 'existingExterior', 'stepProtocol', 'parasolidVersion', 'targetVersion', 'parasolidType', 'outputName'];
let state = freshDefaults(), reviewOpen = false, toastTimer;

function element(tag, text, className) {
  const el = document.createElement(tag);
  if (text !== undefined) el.textContent = text;
  if (className) el.className = className;
  return el;
}
function applyState(c) {
  for (const key of strings) $(key).value = c[key];
  for (const key of checkboxes) $(key).checked = c[key];
  for (const name of ['geometryMode', 'outputFormat']) {
    for (const control of form.querySelectorAll('input[name="' + name + '"]')) control.checked = control.value === c[name];
  }
  for (const control of form.querySelectorAll('input[name="excludedCategories"]')) control.checked = c.excludedCategories.includes(control.value);
}
function readState() {
  const c = freshDefaults();
  for (const key of strings) c[key] = $(key).value;
  for (const key of checkboxes) c[key] = $(key).checked;
  for (const name of ['geometryMode', 'outputFormat']) c[name] = form.querySelector('input[name="' + name + '"]:checked')?.value || c[name];
  c.excludedCategories = [...form.querySelectorAll('input[name="excludedCategories"]:checked')].map(x => x.value);
  return normalize(c);
}
function saveDraft() {
  try { localStorage.setItem(STORAGE_KEY, JSON.stringify(state)); $('save-note').textContent = 'Draft saved in this browser'; }
  catch { $('save-note').textContent = 'Draft available for this session'; }
}
function renderPairs(target, rows) {
  target.replaceChildren();
  for (const [key, value] of rows) target.append(element('dt', key), element('dd', value));
}
function fileCards(target, files) {
  target.replaceChildren();
  for (const f of files) {
    const row = element('div', undefined, 'summary-file');
    row.append(element('span', f.format === 'STEP' ? 'STP' : f.filename.endsWith('.x_b') ? 'X_B' : 'X_T', 'file-icon'));
    const desc = element('div'); desc.append(element('strong', f.filename), element('small', f.compatibility));
    row.append(desc); target.append(row);
  }
}
function renderControls() {
  const reuse = state.existingExterior === 'existing', solid = state.geometryMode === 'solid';
  $('scope-field').hidden = state.assemblyScope === 'entire';
  $('scopeSelection').placeholder = state.assemblyScope === 'components' ? 'AXLE_ASSEMBLY/HOUSING_1\nAXLE_ASSEMBLY/COVER_1' : 'AXLE_ASSEMBLY/CARRIER_1';
  $('exclusion-fields').disabled = !state.excludeInternal;
  $('geometry-fields').disabled = reuse;
  $('existing-note').hidden = !reuse;
  $('sewSurfaces').checked = state.sewSurfaces;
  $('attemptSolid').checked = state.attemptSolid;
  $('sewSurfaces').disabled = solid;
  $('attemptSolid').disabled = solid || !state.sewSurfaces;
  $('solid-help').textContent = solid ? 'Required for a solid target. Retain sheets if the volume is not closed.' : state.sewSurfaces ? 'Requires sewing and a closed volume.' : 'Enable sewing to attempt solid creation.';
  $('quality-help').textContent = { standard: 'A quicker preparation target with less exterior detail.', fine: 'Fine balances exterior detail and processing time.', very_fine: 'Maximum detail target; expect longer processing and larger output.' }[state.quality];
  $('step-field').hidden = state.outputFormat === 'parasolid';
  $('parasolid-field').hidden = state.outputFormat === 'step';
  $('encoding-field').hidden = state.outputFormat === 'step';
  $('target-version-field').hidden = state.outputFormat === 'step' || state.parasolidVersion !== 'custom';
  $('diagram-internals').style.opacity = state.excludeInternal ? '0' : '.8';
}
function renderSummary() {
  const reuse = state.existingExterior === 'existing';
  $('summary-mode').textContent = reuse ? 'Existing result' : state.geometryMode === 'solid' ? 'Solid target' : 'Exterior surfaces';
  $('summary-quality').textContent = reuse ? 'Not applied' : LABELS.quality[state.quality];
  const exclusionCount = state.excludeInternal ? state.excludedCategories.length + state.customExclusions.split(/\r?\n/).filter(x => x.trim()).length : 0;
  renderPairs($('live-summary'), [
    ['Scope', LABELS.scope[state.assemblyScope]],
    ['Visibility', state.visibleOnly ? 'Visible only' : 'All in scope'],
    ['Exclusions', exclusionCount ? exclusionCount + ' selection rules' : 'None'],
    ['Close openings', reuse ? 'Not applied' : state.closeOpenings ? 'Attempt closure' : 'Off'],
    ['Sew surfaces', reuse ? 'Not applied' : state.sewSurfaces ? 'Enabled' : 'Off'],
    ['Create solid', reuse ? 'Not applied' : state.attemptSolid ? 'If closed' : 'Off']
  ]);
  const files = outputs(state);
  fileCards($('summary-files'), files);
  $('file-previews').replaceChildren(...files.map(f => element('span', f.filename, 'file-tag')));
}
function renderReview() {
  renderPairs($('review-summary'), summary(state));
  const reuse = state.existingExterior === 'existing';
  const rows = [
    ['Included components', 'Occurrence paths, visibility and inclusion status · pending NX scan'],
    ['Excluded components', 'Occurrence paths and reasons for each exclusion · pending NX scan'],
    ['Geometry result', 'Actual sheet / solid counts, closure status and single-body outcome · pending NX run'],
    ['Preparation', reuse ? 'Existing geometry will be inspected; no extraction, closure, sewing or solid creation' : 'Extraction, closure and sewing outcomes; unresolved openings and geometry failures'],
    ['Source visibility', state.previewIsolation ? 'Confirm original visibility is restored after preview' : 'Preview isolation is off'],
    ...outputs(state).map(f => [f.format + ' export', f.filename + ' · ' + f.compatibility + ' · planned, not written'])
  ];
  $('report-rows').replaceChildren();
  for (const [key, value] of rows) { const row = element('tr'); row.append(element('td', key), element('td', value)); $('report-rows').append(row); }
  $('review-warnings').replaceChildren(...warnings(state).map(w => element('li', w)));
}
function refresh() { renderControls(); renderSummary(); if (reviewOpen) renderReview(); }
function showErrors(errors) {
  $('form-errors').hidden = !errors.length;
  $('form-errors').replaceChildren(...errors.map(e => element('p', e)));
  for (const id of ['assemblyName', 'scopeSelection', 'outputName', 'targetVersion']) $(id).removeAttribute('aria-invalid');
  if (!errors.length) return;
  if (!state.assemblyName.trim()) $('assemblyName').setAttribute('aria-invalid', 'true');
  if (!state.scopeSelection.trim() && state.assemblyScope !== 'entire') $('scopeSelection').setAttribute('aria-invalid', 'true');
  if (!/^[a-zA-Z0-9][a-zA-Z0-9_-]{0,79}$/.test(state.outputName)) $('outputName').setAttribute('aria-invalid', 'true');
  if (state.outputFormat !== 'step' && state.parasolidVersion === 'custom' && !/^\d{1,2}(\.\d{1,2})?$/.test(state.targetVersion.trim())) $('targetVersion').setAttribute('aria-invalid', 'true');
}
function switchPanel(review) {
  reviewOpen = review;
  $('configure-panel').hidden = review;
  $('review-panel').hidden = !review;
  $('configure-tab').classList.toggle('active', !review);
  $('review-tab').classList.toggle('active', review);
  (review ? $('configure-tab') : $('review-tab')).removeAttribute('aria-current');
  (review ? $('review-tab') : $('configure-tab')).setAttribute('aria-current', 'step');
  if (review) renderReview();
  document.querySelector('.step-nav').scrollIntoView({ behavior: 'smooth', block: 'start' });
}
function goReview() {
  state = readState();
  const errors = validate(state);
  showErrors(errors);
  if (errors.length) { switchPanel(false); $('form-errors').scrollIntoView({ behavior: 'smooth', block: 'center' }); $('form-errors').setAttribute('tabindex', '-1'); $('form-errors').focus({ preventScroll: true }); return; }
  switchPanel(true); $('review-heading').setAttribute('tabindex', '-1'); $('review-heading').focus({ preventScroll: true });
}
function toast(message) {
  $('toast').textContent = message; $('toast').hidden = false;
  clearTimeout(toastTimer); toastTimer = setTimeout(() => { $('toast').hidden = true; }, 3200);
}
function download(filename, content, type) {
  const url = URL.createObjectURL(new Blob([content], { type }));
  const a = element('a'); a.href = url; a.download = filename; document.body.append(a); a.click(); a.remove();
  setTimeout(() => URL.revokeObjectURL(url), 1000);
  $('download-status').textContent = 'Download prepared: ' + filename + '. No CAD geometry has been exported.';
}
function reportText(c) {
  const line = '\n';
  return ['NX 2306 EXTERIOR EXPORT — CONFIGURATION & EXPECTED REPORT', 'Status: NOT EXECUTED. No CAD files have been generated.', '', 'CONFIGURATION', ...summary(c).map(([k, v]) => k + ': ' + v), '', 'EXPECTED RESULT REPORT', 'Included component count and occurrence list: pending NX scan', 'Excluded component count, occurrence list and reasons: pending NX scan', 'Body count: pending NX run; target 1', 'Solid / sheet counts and closure status: pending NX run', 'Source visibility restored: pending NX preview', '', 'PLANNED OUTPUT FILES', ...outputs(c).map(f => f.filename + ' | ' + f.compatibility + ' | planned, not written'), '', 'ITEMS TO VERIFY', ...warnings(c).map(x => '- ' + x), '', 'This is a configuration plan. A compatible NX 2306 workflow must map these options to supported NX operations, inspect the assembly, and produce the actual geometry and result report.', ''].join(line);
}

form.addEventListener('input', () => {
  state = readState(); refresh(); saveDraft();
  if (!$('form-errors').hidden) showErrors(validate(state));
});
form.addEventListener('submit', event => { event.preventDefault(); goReview(); });
$('configure-tab').addEventListener('click', () => switchPanel(false));
$('review-tab').addEventListener('click', goReview);
$('side-review').addEventListener('click', goReview);
$('edit').addEventListener('click', () => switchPanel(false));
$('reset').addEventListener('click', () => { state = freshDefaults(); applyState(state); refresh(); saveDraft(); showErrors([]); toast('Default exterior-export configuration restored.'); });
$('download-config').addEventListener('click', () => {
  try { const payload = packageConfig(state); download(state.outputName + '_configuration.json', JSON.stringify(payload, null, 2) + '\n', 'application/json'); }
  catch (e) { toast(e.message); }
});
$('download-report').addEventListener('click', () => {
  if (validate(state).length) { goReview(); return; }
  download(state.outputName + '_expected_report.txt', reportText(state), 'text/plain;charset=utf-8');
});
$('help').addEventListener('click', () => $('help-dialog').showModal());
$('close-help').addEventListener('click', () => $('help-dialog').close());
$('help-dialog').addEventListener('click', e => { if (e.target === $('help-dialog')) { const r = e.target.getBoundingClientRect(); if (e.clientX < r.left || e.clientX > r.right || e.clientY < r.top || e.clientY > r.bottom) e.target.close(); } });
try {
  const stored = localStorage.getItem(STORAGE_KEY);
  if (stored) {
    const candidate = normalize(JSON.parse(stored));
    if (strings.every(k => typeof candidate[k] === 'string') && checkboxes.every(k => typeof candidate[k] === 'boolean')) state = candidate;
  }
} catch { /* Storage can be unavailable or contain a damaged draft. */ }
applyState(state); refresh();
