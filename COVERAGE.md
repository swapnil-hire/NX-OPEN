# Capability and phase boundaries

Version 0.1 is a pilot implementation with executable source, not a claim that all five development phases or all master-prompt requirements are complete. Overall engineering approval is deliberately blocked by coverage findings.

| Area | Implemented in this source | Remaining / explicitly unverified |
|---|---|---|
| Input | Current displayed/work part, background native `.prt` reference load | Teamcenter managed selector, PDF/DXF/CGM import/parsing/OCR |
| Scope | All sheets/views enumerated; separate current/reference checked lists; resolved membership filtering | UF view-name resolution can be ambiguous; unresolved objects remain included with UNVERIFIED |
| Dimensions | Base Dimension computed size, measurement type, main/dual text, reference/inspection flags, appended text, precision | Complete manual-override detection, dimensional text-to-geometry consistency, basic-dimension relationship validation |
| Tolerances | Base upper/lower API values, tolerance type, hole/shaft deviation and grade, display style | Deriving limits from every tolerance subtype; fit standards tables; full compound/angular-unit conversion |
| GD&T | Modern edit builder read without Commit; characteristic, frame style/count/order, zone shape, material modifier, projected zone, documented scalar frame fields | Full symbol semantics, indicator lists, business modifiers, standard-specific composite interpretation; extended symbolic text |
| Datum references | Primary/secondary/tertiary roles kept distinct; ordered common-datum members and modifiers captured; missing extracted definitions flagged | Definite absence outside selected scope, datum targets, common-datum precedence, full attachment equivalence |
| Datums | BaseDatum labels, count/repetition diagnostics, attachment-count evidence | Repetition can be legitimate; no automatic duplicate-feature judgement |
| Notes | Raw NX text and whitespace-only normalized comparison; optional approved vocabulary warnings | Material/process/torque/thread language parsing; synonym equivalence; spelling certification |
| Symbols | Inventory of supported Annotation wrappers; raw SimpleDraftingAid text where exposed | Weld/surface texture/finish/custom symbol structure and technical interpretation |
| Views | Subtype name, scale, angle, orientation matrix, origin, broken/out-of-date status, hidden-line and label toggles | Section arrows, labels' complete contents, alignment/broken/partial boundaries and geometric correspondence |
| Title block | Explicit user-configured string part-attribute mappings and expected values | Displayed title-block/table-cell content, attribute-link freshness, Teamcenter integration, revision history |
| Standards | Explicit allowed sheet dimensions, projection enum and configured expected string values | No automatic ASME Y14.5 / ISO GPS rules, interpretations or certification |
| Matching | Common CHECK_ID; optional trusted journal IDs; mutual unique name/content candidates | Geometry-based matching; moved/recreated objects without stable identity; ambiguous repeated annotations |
| UI | Setup, categories, scope, progress/cancel at scan/match boundaries, status filter, values/diagnostics | Block Styler integration, confidence approval/pair-mapping editor, reference-object navigation |
| Graphics | Native temporary current-object highlight and clear; grid severity colours | Multi-colour NX overlays, automatic zoom/open-sheet (view settings intentionally not changed), reference highlighting |
| Images | Current active graphics PNG; supplied reference image embedded in Excel | All-sheet snapshots, registered difference markers, graphical image comparison |
| Report | Five requested workbook sheets plus two raw inventories; CSV, logs, unique run folders; private Excel instance cleanup | Automatic PDF report generation and XLSX without Excel |
| Validation | Local source static checks performed; source includes VB self-tests and local compile helper | Compilation against installed NX 2306 DLLs, actual drafting fixtures, Excel COM run and preservation audit |

## Match interpretation

A field PASS means the extracted field values agree under the reported correspondence and comparison rule. It never proves that the same geometric feature was inspected. The matching method is retained in each finding. Names/content require correspondence review; ambiguous candidates remain unresolved.

Exact mode treats text/formatting field differences as FAIL. Engineering mode still compares nominal/tolerance numbers and structural enums independently, while format-only differences are warnings. Whitespace normalization preserves case, all numeric tokens, datum order, punctuation and NX symbol codes. It does not translate symbols or recognize equivalent manufacturing language. FCF scalar strings that contain expressions or different textual number formats are conservatively compared as text in this version.

No tolerance is made less strict to obtain a match. `NumericEpsilon` only addresses numerical representation; it is capped in the UI at 0.001 native units. Mixed part units prevent a numeric PASS; unit conversion and angular/native-unit interpretation require local fixtures.

## Preservation contract

No source drawing Save, Close, Update, annotation edit, component edit, attribute write or colour assignment occurs in this source. Edit builders are queried and destroyed without Commit. ImageExportBuilder.Commit writes an external PNG only. The journal does not change sheets, view orientation, zoom or camera settings. Normal NX transient highlighting is cleared on command, form close and module exit. Avoid running with an existing manual highlight that you need preserved; this version does not query pre-existing highlight state.

Background opening leaves the reference loaded; any source already modified before the tool runs remains the user's responsibility. The tool does not claim to suppress NX's own load-time migrations or all vendor-internal read side effects; compare dirty/save status during local validation.

## Next implementation stages

1. Validate this extractor against a hand-counted fixture set, capture subtype failures, confirm unit semantics and ownership, then add override and table-cell readers.
2. Add a user-approved mapping editor and geometric parent signatures for related revisions. Keep ambiguous matches distinct from absences.
3. Add a recorded/verified read-only navigation adapter that restores sheet/view settings; implement transient category overlays without recolouring source objects.
4. Implement separate parser adapters for searchable PDF and DXF with units/symbol/font provenance. CGM requires a validated converter/parser. OCR output remains uncertain until checked against engineering symbols and values.
5. Use the production matrix and configuration-controlled standards rules with exact edition/clause references and engineering sign-off. Only then consider changing the release-coverage gate.
