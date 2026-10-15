# Local validation and acceptance matrix

## Results actually obtained during development

- Python source consistency checks passed: VB strings, delimiters, routine/block counts, builder cleanup and prohibited source mutation scan.
- Built-in comparison-engine VB tests are included, with a separate compile/run PowerShell helper. They were not executed here because no VB/.NET runtime or compiler was installed.
- No NX 2306 execution, licensed builder call, Excel COM export, drawing fixture comparison or source-preservation audit was executed here.
- The sample workbook contains synthetic values and is not an NX run report.

## Before the first pilot

Use copied non-production drawings and a separate output folder. Record NX release/build and NXOpen DLL versions. Ensure all required drafting/model dependencies are loaded. Hand-count supported objects by sheet and type. Save copies before the test and record the file hashes, dirty state, active sheet, view orientation/scale and annotation formatting. Do not save either source from the journal.

First run the optional engine tests and compile helper in Windows PowerShell 5.1. Then run **Extract current** inside NX and compare inventory values with NX's Edit/Information dialogs. Resolve extraction errors before evaluating comparison accuracy.

| Fixture / deliberate change | Expected evidence / acceptance |
|---|---|
| Two identical related copies with common unique CHECK_ID values | Selected field values agree; required uncovered checks still prevent overall release PASS |
| Diameter 25 mm, symmetric tolerance 0.05 → 0.10 | Upper/lower API tolerance FAIL; nominal 25 remains PASS; no averaging or tolerance relaxation |
| Nominal 25 → 26 with matching displayed override kept at 25 | Computed nominal discrepancy must be visible; override verification remains UNVERIFIED until adapter validated |
| Equal nominal text but different basic/reference/inspection/tolerance type | Relevant extracted flags/enums differ; verify basic representation in tolerance-type API |
| Linear, angular, radius and diameter each with unilateral/bilateral/limit/fit modes | Hand-check raw API values, units, signs and displayed text; no inferred min/max until subtype semantics validated |
| Inch part vs metric reference | Numeric comparison UNVERIFIED rather than a false equality/conversion |
| Hole/thread, chamfer, ordinate/baseline/chain groups | Base fields/text captured; missing compound/subtype semantics explicitly UNVERIFIED |
| FCF characteristic position → concentricity/profile/etc. | Characteristic field FAIL, no visual-symbol substitution |
| FCF tolerance 0.20 → 0.30, same CHECK_ID | Structured tolerance text FAIL; original/current raw values retained |
| FCF material modifier MMC → LMC / RFS representation | MaterialModifier enum difference visible; no equivalent-symbol assumption |
| Datum sequence A/B/C → A/C/B | Secondary/Tertiary field FAIL |
| Composite vs single-segment / multiple segment changes | FrameStyle or frame count/order changes; extended semantic verification remains UNVERIFIED |
| Projected zone or common datum changed | Captured flags/value/member order differences; extended datum meaning remains UNVERIFIED |
| Datum definition missing, renamed, repeated on another view | Definition unresolved or repeated warning; out-of-scope definitions never labelled definitively missing |
| Repeated identical dimensions without CHECK_ID | Ambiguous WARNING; no arbitrary pairing or false PASS |
| Unique object names with different geometric references | Correspondence review warning; attachment evidence visible, no geometric-equivalence assertion |
| Same engineering content moved across sheets / view rearrangement | Field agreement plus layout/context warning; no engineering mismatch merely from movement |
| Native object deleted and no remaining same-kind candidates | MISSING within extracted scope; current-only counterpart becomes EXTRA |
| Multiple plausible unmatched same-kind objects | Ambiguous reference/current warnings instead of guessed MISSING/EXTRA |
| Notes with 0.5 → 5, 1045 → 1050, A-B → A/B | Text FAIL; normalization must preserve technical distinctions |
| Notes differ only by whitespace | Engineering mode WARNING; Exact mode FAIL |
| Unknown spelling words / names / abbreviations | Vocabulary review WARNING, no auto-correction |
| A0–A4/custom sheet, first/third angle, title attributes | Extracted metadata; only explicit company-profile rules execute |
| Title attributes disagree with displayed table cells | Report notes title-cell verification is not implemented; no assurance from attributes alone |
| Weld/surface finish/custom/table/PMI/inherited annotation | Unsupported structures inventoried/diagnosed; no silent coverage PASS |
| Selected current/reference sheets and views differ | Effective scope recorded; unresolved ownership retained with UNVERIFIED |
| Part load warnings, missing dependencies, stale views | UNVERIFIED/WARNING/ERROR retained and overall status cannot be release PASS |
| Select result then Highlight/Clear/close/cancel/error | Only tool-created transient highlight cleared; formatting/colour/content unchanged |
| Cancel during traversal/matching | No completed inspection/export presented; source unchanged |
| Excel installed / absent / SaveAs failure / missing image | XLSX success or clear error, CSV retained; private Excel workbook closes and instance quits |
| Report contains =1+1 or +formula-like note | Excel text remains literal; CSV prefix prevents formula execution |
| Long note over Excel cell limit | Workbook visibly marks truncation; CSV retains full text |
| Repeat scans / repeated Excel exports | Unique output folder/workbook names; no existing files/source overwritten |

## Preservation audit

After every runtime test, clear the tool highlight. Compare object counts, text, tolerances, layers, colours, line settings, ownership, associativities, view state, active sheet and part dirty state against the baseline. Source files' disk hashes must remain unchanged. The reference remains loaded intentionally; close it manually if desired. If merely reading an edit builder dirties a part in your NX build, block that adapter and record UNVERIFIED until replaced with a verified reader.

## Production gate

Do not remove the unconditional release-coverage findings merely to obtain PASS. Production acceptance needs complete category/subtype coverage, matching adjudication, validated standards edition/rule configuration, no unresolved required checks, traceable test evidence and engineering sign-off. PDF/DXF/CGM adapters require their own provenance, units, font/symbol and OCR uncertainty tests.
