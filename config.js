export const DEFAULTS = Object.freeze({
  assemblyName: 'AXLE_ASSEMBLY', assemblyScope: 'entire', scopeSelection: '',
  visibleOnly: true, excludeInternal: true, excludedCategories: ['gears', 'bearings', 'seals', 'other'], customExclusions: '',
  geometryMode: 'surface', quality: 'fine', closeOpenings: true, sewSurfaces: true, attemptSolid: false,
  existingExterior: 'extract', previewIsolation: true,
  outputFormat: 'both', stepProtocol: 'AP242', parasolidVersion: 'native', targetVersion: '', parasolidType: 'text', outputName: 'axle_exterior'
});
export const LABELS = {
  scope: { entire: 'Entire assembly', subassembly: 'Selected subassembly', components: 'Selected components' },
  quality: { standard: 'Standard', fine: 'Fine', very_fine: 'Very fine' },
  categories: { gears: 'Internal gears', bearings: 'Internal bearings', fasteners: 'Internal fasteners', seals: 'Internal seals', other: 'Other internal components' }
};
export function freshDefaults() { return { ...DEFAULTS, excludedCategories: [...DEFAULTS.excludedCategories] }; }
export function lines(value) { return String(value || '').split(/\r?\n/).map(x => x.trim()).filter(Boolean); }
export function normalize(raw) {
  const c = { ...freshDefaults(), ...raw };
  c.excludedCategories = Array.isArray(c.excludedCategories) ? c.excludedCategories.filter(x => x in LABELS.categories) : [];
  if (c.geometryMode === 'solid') { c.sewSurfaces = true; c.attemptSolid = true; }
  if (!c.sewSurfaces) c.attemptSolid = false;
  return c;
}
export function validate(c) {
  const errors = [];
  const choices = { assemblyScope: ['entire', 'subassembly', 'components'], geometryMode: ['surface', 'solid'], quality: ['standard', 'fine', 'very_fine'], existingExterior: ['extract', 'existing'], outputFormat: ['step', 'parasolid', 'both'], stepProtocol: ['AP214', 'AP242'], parasolidVersion: ['native', 'custom'], parasolidType: ['text', 'binary'] };
  for (const [field, valid] of Object.entries(choices)) if (!valid.includes(c[field])) errors.push('Choose a valid ' + field + '.');
  if (!c.assemblyName.trim()) errors.push('Enter an assembly reference.');
  if (!/^[a-zA-Z0-9][a-zA-Z0-9_-]{0,79}$/.test(c.outputName)) errors.push('Use 1–80 letters, numbers, hyphens or underscores for the output name. Start with a letter or number.');
  if (c.assemblyScope !== 'entire' && !lines(c.scopeSelection).length) errors.push('Enter the component occurrence path(s) for the selected scope.');
  if (c.assemblyScope === 'subassembly' && lines(c.scopeSelection).length > 1) errors.push('Use one occurrence path for a selected subassembly.');
  if (c.outputFormat !== 'step' && c.parasolidVersion === 'custom' && !/^\d{1,2}(\.\d{1,2})?$/.test(c.targetVersion.trim())) errors.push('Enter a Parasolid target version such as 32.0.');
  return errors;
}
export function outputs(c) {
  const files = [];
  if (c.outputFormat !== 'parasolid') files.push({ format: 'STEP', compatibility: c.stepProtocol, filename: c.outputName + '.stp' });
  if (c.outputFormat !== 'step') files.push({ format: 'Parasolid', compatibility: c.parasolidVersion === 'native' ? 'NX 2306 native translator version' : 'Target ' + c.targetVersion.trim() + ' · verify translator support', filename: c.outputName + (c.parasolidType === 'binary' ? '.x_b' : '.x_t') });
  return files;
}
export function warnings(c) {
  const w = [];
  if (c.existingExterior === 'existing') w.push('Existing exterior: extraction, closure, sewing and solid creation settings will not be applied. The existing result must be inspected in NX.');
  else {
    if (c.attemptSolid) w.push('A solid is possible only if the sewn surfaces form a closed volume. Keep sheets and flag review if solid creation fails.');
    if (!c.sewSurfaces) w.push('Sewing is off. The result can contain multiple sheet bodies.');
    if (!c.closeOpenings) w.push('Openings will be retained; the result may remain open.');
    if (c.quality === 'very_fine') w.push('Very fine extraction can increase processing time and geometry size.');
  }
  w.push('The single-body target depends on connected geometry. Report the actual body count and flag any result that is not one body.');
  if (c.outputFormat !== 'step' && c.parasolidVersion === 'custom') w.push('Confirm the requested Parasolid version is supported by the installed NX translator and the recipient.');
  return w;
}
export function summary(c) {
  const reuse = c.existingExterior === 'existing';
  const exclusions = c.excludeInternal ? [...c.excludedCategories.map(x => LABELS.categories[x]), ...lines(c.customExclusions)] : [];
  return [
    ['Assembly', c.assemblyName.trim()],
    ['Scope', LABELS.scope[c.assemblyScope] + (c.assemblyScope === 'entire' ? '' : ' · ' + lines(c.scopeSelection).join(', '))],
    ['Visibility', c.visibleOnly ? 'Visible components only' : 'All components in scope'],
    ['Internal exclusions', exclusions.length ? exclusions.join(', ') : 'None selected'],
    ['Source', reuse ? 'Export existing exterior' : 'Extract a new exterior in the existing part'],
    ['Geometry', reuse ? 'Use existing geometry as-is' : c.geometryMode === 'solid' ? 'Solid target · sheet fallback' : 'Exterior surfaces'],
    ['Quality', reuse ? 'Not applied' : LABELS.quality[c.quality]],
    ['Close openings', reuse ? 'Not applied' : c.closeOpenings ? 'Attempt closure' : 'Keep openings'],
    ['Sew surfaces', reuse ? 'Not applied' : c.sewSurfaces ? 'Attempt to join adjacent sheets' : 'Off'],
    ['Solid creation', reuse ? 'Not applied' : c.attemptSolid ? 'Attempt only for a closed volume' : 'Off'],
    ['Preview isolation', c.previewIsolation ? 'Hide sources temporarily, then restore' : 'Off'],
    ['Result target', 'One unassociated body in the existing part'],
    ['Output', outputs(c).map(x => x.format + ' · ' + x.compatibility).join(' + ')]
  ];
}
export function expectedReport(c) {
  return {
    reportType: 'expected_pre_export_report', executionStatus: 'not_run',
    assemblyReference: c.assemblyName.trim(),
    includedComponents: { count: null, occurrences: [], status: 'Determined in NX from scope, visibility and exclusions' },
    excludedComponents: { count: null, occurrences: [], reasons: [], status: 'Determined in NX; record each excluded occurrence and reason' },
    geometry: { bodyCount: null, sheetBodyCount: null, solidBodyCount: null, targetBodyCount: 1, closedVolume: null, status: 'Not measured; flag disconnected geometry, failed closure or body count other than one' },
    sourceVisibilityRestored: null,
    exportFiles: outputs(c).map(x => ({ ...x, status: 'planned', written: false })),
    warnings: warnings(c), errors: []
  };
}
export function packageConfig(raw) {
  const c = normalize(raw), errors = validate(c);
  if (errors.length) throw new Error(errors.join('\n'));
  const reuse = c.existingExterior === 'existing';
  return {
    schema: 'nx2306.exterior-export.configuration.v1', nxRelease: '2306', executionStatus: 'not_run',
    workflow: { destination: 'existing_part', associativity: false, targetBodyCount: 1, failureBehavior: 'retain_sheet_result_and_report', configurationOnly: true },
    assembly: { reference: c.assemblyName.trim(), scope: c.assemblyScope, selectedOccurrencePaths: c.assemblyScope === 'entire' ? [] : lines(c.scopeSelection), visibleOnly: c.visibleOnly, exclusions: { enabled: c.excludeInternal, categories: c.excludeInternal ? c.excludedCategories : [], explicitOccurrencePaths: c.excludeInternal ? lines(c.customExclusions) : [] } },
    geometry: { source: c.existingExterior, mode: reuse ? 'existing' : c.geometryMode, quality: reuse ? null : c.quality, closeOpenings: reuse ? false : c.closeOpenings, sewSurfaces: reuse ? false : c.sewSurfaces, attemptSolidWhenClosed: reuse ? false : c.attemptSolid, previewIsolation: c.previewIsolation },
    export: { format: c.outputFormat, outputBaseName: c.outputName, stepProtocol: c.outputFormat === 'parasolid' ? null : c.stepProtocol, parasolid: c.outputFormat === 'step' ? null : { version: c.parasolidVersion === 'native' ? 'nx_native' : c.targetVersion.trim(), encoding: c.parasolidType, verifyTranslatorSupportAtRuntime: true }, files: outputs(c) },
    expectedResultReport: expectedReport(c)
  };
}
