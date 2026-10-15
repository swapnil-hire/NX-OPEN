import { test } from 'node:test';
import assert from 'node:assert/strict';
import { readFileSync } from 'node:fs';
import { freshDefaults, normalize, validate, outputs, warnings, packageConfig, expectedReport, summary } from '../dist/config.js';

test('default plan includes both formats and pending counts; it requests no new part or associativity', () => {
  const c = freshDefaults(), p = packageConfig(c);
  assert.deepEqual(validate(c), []);
  assert.deepEqual(outputs(c).map(x => x.filename), ['axle_exterior.stp', 'axle_exterior.x_t']);
  assert.equal(p.export.stepProtocol, 'AP242');
  assert.equal(p.workflow.destination, 'existing_part');
  assert.equal(p.workflow.associativity, false);
  assert.equal(p.workflow.targetBodyCount, 1);
  assert.equal(p.expectedResultReport.geometry.bodyCount, null);
  assert.ok(p.expectedResultReport.exportFiles.every(f => f.written === false));
});
test('solid target requires sewing and preserves sheet fallback', () => {
  const c = normalize({ geometryMode: 'solid', sewSurfaces: false, attemptSolid: false });
  assert.equal(c.sewSurfaces, true); assert.equal(c.attemptSolid, true);
  assert.match(packageConfig(c).workflow.failureBehavior, /retain_sheet/);
  assert.ok(warnings(c).some(w => /closed volume/.test(w)));
});
test('surface mode can disable sewing and solid creation together', () => {
  const c = normalize({ geometryMode: 'surface', sewSurfaces: false, attemptSolid: true });
  assert.equal(c.attemptSolid, false);
  assert.ok(warnings(c).some(w => /multiple sheet/.test(w)));
});
test('selected scope requires occurrence paths and subassembly is singular', () => {
  assert.ok(validate(normalize({ assemblyScope: 'components' })).length);
  assert.deepEqual(validate(normalize({ assemblyScope: 'components', scopeSelection: 'ROOT/COVER_1\nROOT/COVER_2' })), []);
  assert.ok(validate(normalize({ assemblyScope: 'subassembly', scopeSelection: 'ROOT/A\nROOT/B' })).length);
  const p = packageConfig(normalize({ assemblyScope: 'components', scopeSelection: ' ROOT/COVER_1\nROOT/COVER_2 ' }));
  assert.deepEqual(p.assembly.selectedOccurrencePaths, ['ROOT/COVER_1', 'ROOT/COVER_2']);
});
test('internal exclusion rules and explicit paths can be fully switched off', () => {
  const c = normalize({ excludeInternal: false, customExclusions: 'ROOT/GEAR_1', visibleOnly: false });
  const p = packageConfig(c); assert.equal(p.assembly.visibleOnly, false);
  assert.deepEqual(p.assembly.exclusions.categories, []); assert.deepEqual(p.assembly.exclusions.explicitOccurrencePaths, []);
  assert.ok(summary(c).some(([k, v]) => k === 'Internal exclusions' && v === 'None selected'));
});
test('existing geometry bypasses extraction, closure, sewing and solid creation', () => {
  const c = normalize({ existingExterior: 'existing', geometryMode: 'solid' });
  const p = packageConfig(c);
  assert.deepEqual([p.geometry.mode, p.geometry.quality, p.geometry.closeOpenings, p.geometry.sewSurfaces, p.geometry.attemptSolidWhenClosed], ['existing', null, false, false, false]);
  assert.ok(warnings(c).some(w => /will not be applied/.test(w)));
});
test('STEP-only AP214 ignores inactive Parasolid fields', () => {
  const c = normalize({ outputFormat: 'step', stepProtocol: 'AP214', parasolidVersion: 'custom', targetVersion: '' });
  assert.deepEqual(validate(c), []); assert.equal(outputs(c).length, 1);
  assert.equal(packageConfig(c).export.parasolid, null); assert.equal(packageConfig(c).export.stepProtocol, 'AP214');
});
test('Parasolid encoding and requested compatibility appear in plan and require validation', () => {
  const invalid = normalize({ outputFormat: 'parasolid', parasolidVersion: 'custom', targetVersion: 'bad' });
  assert.ok(validate(invalid).length); assert.throws(() => packageConfig(invalid));
  const valid = normalize({ ...invalid, targetVersion: '32.0', parasolidType: 'binary' });
  assert.equal(outputs(valid)[0].filename, 'axle_exterior.x_b');
  assert.equal(packageConfig(valid).export.stepProtocol, null);
  assert.equal(packageConfig(valid).export.parasolid.version, '32.0');
  assert.equal(packageConfig(valid).export.parasolid.verifyTranslatorSupportAtRuntime, true);
});
test('unsafe filenames and unknown options cannot be packaged', () => {
  for (const name of ['../export', 'file.stp', 'name with space', '', 'a'.repeat(81)]) assert.ok(validate(normalize({ outputName: name })).length);
  assert.ok(validate(normalize({ outputFormat: 'made-up' })).length);
  assert.deepEqual(validate(normalize({ outputName: 'AXLE-186_v2' })), []);
});
test('all output combinations produce planned reports, never fabricated execution results', () => {
  for (const outputFormat of ['step', 'parasolid', 'both']) for (const geometryMode of ['surface', 'solid']) for (const existingExterior of ['extract', 'existing']) {
    const c = normalize({ outputFormat, geometryMode, existingExterior }), report = expectedReport(c);
    assert.equal(report.executionStatus, 'not_run'); assert.equal(report.includedComponents.count, null); assert.equal(report.excludedComponents.count, null);
    assert.equal(report.geometry.bodyCount, null); assert.equal(report.geometry.closedVolume, null);
    assert.equal(report.exportFiles.length, outputFormat === 'both' ? 2 : 1);
    assert.ok(report.exportFiles.every(f => f.status === 'planned' && !f.written));
  }
});
test('page has every application target and unique IDs; all assets are local', () => {
  const html = readFileSync(new URL('../dist/index.html', import.meta.url), 'utf8');
  const app = readFileSync(new URL('../dist/app.js', import.meta.url), 'utf8');
  const ids = [...html.matchAll(/\bid="([^"]+)"/g)].map(m => m[1]);
  assert.equal(ids.length, new Set(ids).size);
  for (const match of app.matchAll(/\$\('([^']+)'\)/g)) assert.ok(ids.includes(match[1]), 'Missing HTML id: ' + match[1]);
  assert.ok(!/\b(?:src|href)="https?:/.test(html));
  assert.ok(!/\bfetch\s*\(|XMLHttpRequest/.test(app));
  assert.ok(html.includes('no NX connection'));
});
