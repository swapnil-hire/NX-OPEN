# Siemens NX 2306: Automated Drawing Checker — Phase 1/2 engineering prototype

**Status:** Generated against published NXOpen SC2306 API references, **not compiled or executed in a live NX 2306 installation.** Requires in-NX validation before relying on results. This is a real, single-file VB.NET NX journal implementation, not pseudocode, but it is **NOT** a complete ASME Y14.5/ISO GPS compliance engine.

## Files
- `NX2306_Drawing_Compare_Phase1_2.vb` — journal with selection dialog, extraction, native `.prt` reference compare, Excel/CSV export, temporary highlighting.
- `README_NX2306_Drawing_Compare.md` — instructions and validation.

## Start in NX 2306
1. Make backup copies of two known drawing `.prt` files; ensure no uncommitted changes in the source.
2. Display the CURRENT drawing. Open the Drafting application and an existing sheet. For the initial test use local file-based parts, not Teamcenter identifiers.
3. Go to **Tools → Journal → Play** and choose the `.vb` source.
4. In the dialog, browse to REFERENCE `.prt` or leave blank for extraction only.
5. Choose an `.xlsx` output name and check categories. Choose a distance for **unique position pairing** (default 15 mm) appropriate to drawing layout. The distance is in drawing part model-space millimeters after conversion, NOT pixel distance.
6. Optionally capture the currently shown NX graphics view. That is not a per-sheet image comparison.
7. Read the NX Listing Window; inspect **Inspection Summary**, **Detailed Comparison**, **Missing Modified**, **Drawing Images**, and **Validation Log**.
8. If selected, use the findings dialog to temporarily highlight an object on the displayed sheet; close or clear to remove highlighting.
9. Do **not** approve drawings automatically based on this prototype.

## Dependency / software
- Siemens NX 2306 on Windows with Drafting license/appropriate availability.
- Built-in NXOpen.dll, NXOpen.UF.dll and NXOpen.Utilities.dll assemblies.
- .NET Windows Forms and System.Runtime.InteropServices.
- Local Microsoft Excel desktop required for `.xlsx` writing by COM late binding (no additional Excel Interop reference is required). If Excel is absent or export fails, the journal creates `<basename>_comparison.csv` and `<basename>_warnings.csv` as fallback. The screenshot PNG (if created) is separate. CSV cannot embed images or workbook formatting.
- The reference `.prt` is loaded via `Session.Parts.OpenBase`, without switching the displayed part. This journal deliberately does not close the reference automatically, because it may have been open beforehand.

## What is implemented in the initial version
- Native `.prt` current + reference selection; local full-file paths.
- Sheet enumeration with names, sizes and projection types.
- Drafting-view enumeration with names, .NET runtime view type, scale and drawing reference point.
- UF object enumeration of dimension (`UF_dimension_type`) and drafting (`UF_drafting_entity_type`) objects.
- Native dimension measurements: computed size, tolerance type, upper/lower tolerance numeric fields, text and secondary text.
- Note/label and GDT encoded text reads using `SimpleDraftingAid.GetText` where supported.
- Simple datum identification based on the returned .NET runtime type name, not an independent semantic datum parser.
- Sheet association attempted through `UFView.AskViewDependentStatus`, `UFView.AskTagOfViewName`, `UFDraw.AskDrawingOfView`. Unresolvable sheet membership is marked **UNVERIFIED**.
- Exact unique string matching within same-named sheets; if different, a **unique** nearby annotation of the same runtime type can be tentatively compared. Multiple candidates are **UNVERIFIED**. It is NOT robust associativity-based matching.
- Explicit PASS/FAIL/MISSING/EXTRA/UNVERIFIED and descriptive reasons.
- Optional temporary `DisplayableObject.Highlight` / `Unhighlight`, only for current drawing object tags. No saved color change.
- Optional `UFDisp.CreateImage` capture of CURRENT visible NX graphics view. **Deprecated** NX API, intentionally wrapped in an exception handler.
- Excel `.xlsx` through COM, and CSV fallback; no CAD part saves.

## Important limitations (not implemented; do not misrepresent)
- Full structural GD&T characteristic interpretation, modifier parsing, segment hierarchy, composite tolerance, projected tolerance, datum precedence, and ASME / ISO conformance rules.
- Associativity-based matching between revisioned drawings. Matching by `Tag` or `JournalIdentifier` across independent part files is not guaranteed. Nearest anchor alone is NOT sufficient for release-level correctness.
- Semantically equivalent GDT notes/ symbols, actual tolerance formula evaluation, basic/reference semantics beyond captured properties, thread/hole callout parsing, plus/minus text normalization, and material/process note equivalence.
- Datum targets, section-arrow integrity, title block table cells, projection-orientation matrices, overridden dimensions' semantic equivalence, drafting symbol inventories (surface finish/weld etc.) and spelling dictionaries.
- Automated comparison with PDF, DXF, CGM; OCR/vector parsing; image overlays, pixel diffs and image capture of all reference sheets.
- Multi-sheet mappings when sheets are renamed or reordered. Sheets must share names, otherwise MISSING/EXTRA are generated and manual mapping is required.
- Auto-sheet switch/zoom-to-object in review; highlight will be invisible when object is on another sheet.
- Excel report of part number/material/revision unless separately derived from title block/attributes; that is Phase 3/4 work.
- Tag collisions during independent part loads, annotation objects returned through inconsistent class hierarchies and note string encodings may require adjustment in the installed NX 2306 system.

### Critical semantics
- **PASS** means the extracted representation agreed; it is not a certificate of engineering correctness.
- **UNVERIFIED** means cannot certify semantic match or object extraction. The overall report is intentionally *INCOMPLETE / UNVERIFIED*, even if simple text comparisons pass, until a broader checker is implemented.
- **Dimension values:** non-angular values converted to mm based on source part units; angular units are preserved. Verify NX convention for each special dimension type (ordinate, degrees-minutes-seconds, inferred/overridden text) before sign-off.
- **Warnings** record unsupported objects and view/sheet association failures. Do not silently ignore them.
- **Example failure**: reference diameter has nominal 25 with +0.05/-0.05 and current has +0.10/-0.10; if unique match by anchor, report FAIL; otherwise report ambiguity, not made-up matching.

## Sample output (illustrative)
| ID | Category | Sheet | Reference | Current | Status | Reason |
|---|---|---|---|---|---|---|
| CHK-00001 | DIMENSION | Sheet1 | Nominal=25;Tol=Bilateral;Upper=0.05;Lower=-0.05 | Nominal=25;Tol=Bilateral;Upper=0.1;Lower=-0.1 | FAIL | Unique nearby object differs |
| CHK-00002 | NOTE | Sheet1 | REMOVE BURRS | REMOVE BURRS | PASS | Exact content match |
| CHK-00003 | GDT | Sheet1 | [encoded position frame] | [encoded position frame] | UNVERIFIED | Raw text matches; semantic GD&T still unverified |
| CHK-00004 | VIEW | Sheet1 | SECTION A-A | (not found) | MISSING | No matching object found |

## Validation procedure
**1. API smoke test.** In NX 2306, run against a local test drawing with one sheet, one linear dimension and one note. Confirm both appear in the Detailed Comparison in extraction-only mode. If compilation fails, check the exact installed NXOpen DLLs and share line number/error; do not suppress it.

**2. Baseline identical drawings.** Copy a drawing to a different `.prt` file name without changing annotations. Compare. Confirm sheet names match, dimension object categories appear, and you see no invented FAIL for unchanged unique annotations. All semantic/GD&T scope still UNVERIFIED.

**3. Deliberate dimension tolerance change.** Change a single unique dimension from ±0.05 to ±0.10 in test drawing; leave annotation near same position. Confirm change is recorded as FAIL with both tolerance values.

**4. Annotation removal test.** Remove one unique note/datum in test copy; confirm MISSING or UNVERIFIED rather than silent PASS.

**5. Ambiguity test.** Place two dimensions of the same type with the same text near the original annotation. Confirm tool says UNVERIFIED (ambiguous) rather than falsely pairing.

**6. Move/reorder test.** Move notes > matching distance, swap sheets, rename sheet. Confirm diagnostic behavior. Improve algorithm in future before production use.

**7. Unit conversion.** Repeat with inch-source part and millimeter-source part; confirm numeric dimensions in mm for linear. Test angular dimensions separately.

**8. Workbook/fallback.** Test on PC with Excel and another without Excel; verify workbook/CSV, formula protection, cleanup, and screenshot behavior.

**9. No-change safety test.** Record NX modified flags before/after; run journal and verify it has not changed geometry, view visibility, annotation display properties or saved the part. Highlights should clear after form closes.

## NXOpen SC2306 documentation reviewed
- https://nxopencsdocumentation.thescriptingengineer.com/SC2306/NXOpen.Annotations.Dimension.html
- https://nxopencsdocumentation.thescriptingengineer.com/SC2306/NXOpen.Annotations.SimpleDraftingAid.html
- https://nxopencsdocumentation.thescriptingengineer.com/SC2306/NXOpen.Drawings.DrawingSheet.html
- https://nxopencsdocumentation.thescriptingengineer.com/SC2306/NXOpen.Drawings.DraftingView.html
- https://nxopencsdocumentation.thescriptingengineer.com/SC2306/NXOpen.UF.UFView.html
- https://nxopencsdocumentation.thescriptingengineer.com/SC2306/NXOpen.UF.UFDrf.html
- https://nxopencsdocumentation.thescriptingengineer.com/SC2306/NXOpen.UF.UFDisp.html
- https://nxopencsdocumentation.thescriptingengineer.com/SC2306/NXOpen.DisplayableObject.html

### Next engineering stages
- Phase 3: robust user UI, view-to-view association, item selection/zoom, per-sheet screenshots and alignment, title blocks/extra symbols.
- Phase 4: PDF/DXF/CGM through a separate explicit parser subsystem, not a fake native CAD data model; image OCR only where necessary.
- Phase 5: validate with real NX 2306 journals and sample drawings, build structured GD&T extractor via `Annotations.Fcf`/`FcfData` and conventional standard profile rules; create regression tests before production release.
