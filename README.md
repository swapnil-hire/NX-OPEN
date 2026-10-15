# NX 2306 native drawing checker — staged version 0.1

This package contains complete VB.NET source for a bounded first implementation: native drawing extraction, conservative comparison, interactive results, temporary highlighting, CSV and Excel reporting. It is not the complete engineering-release application described in the master prompt. Read COVERAGE.md before interpreting results.

No NX 2306 installation, licensed NX session or desktop Excel was available during development. The source has received documentation and static consistency checks; successful compilation and execution in your NX installation are not claimed.

## Start in NX

1. Extract this ZIP into a local folder, for example `C:\NX_Journals\DrawingChecker`.
2. Use **copies** of two fully loaded native NX drawing `.prt` files. Open the current drawing in Drafting and make it both displayed part and work part. Open the sheet you want in the screenshot.
3. Choose **Tools → Journal → Play**, select `NX2306_Drawing_Checker.vb`, and run it. No DLL build is required for the journal route.
4. First select **Extract current**. Review the inventory and diagnostics before attempting comparison.
5. To compare, use **Load .prt** to select the reference. It opens without changing the displayed/work part and remains open when the journal exits. For Teamcenter-managed drawings, use a locally authorized native export for this version; a managed-file selection dialog is not implemented.
6. Select the output folder, categories, and independent current/reference sheet and view scopes. Leave **All** enabled for the first test. Sheet-level notes can be included separately from view filtering. Unresolved membership is retained and flagged rather than silently excluded.
7. Choose **Engineering**, **Exact**, or **Configured company rules**. Numeric equality epsilon is a small floating-point comparison allowance in native API units; it is not permission to increase a drawing tolerance. Start with `0.000001` and 1 mm for layout warnings.
8. Click **Compare**. Filter the results by severity. Select a row and click **Highlight current** to apply normal NX temporary highlighting. Open the reported sheet manually. No automatic sheet opening, view zoom or permanent category colours are applied.
9. Click **Clear highlight**, or close the dialog to clear the highlight created by this tool.
10. Click **Export Excel** to create the `.xlsx` report. Desktop Microsoft Excel is required only for this button; CSV reports are already written after a successful scan.

If a compilation message appears, retain the **complete error text and line number**. Use the compile helper and the API matrix before changing a member name. Do not replace a failed API with an assumed property.

## What is in the report

The workbook has the requested first five sheets plus two raw inventories:

1. Inspection Summary — drawing identities, mapped part number/revision, effective scope, row counts and overall status.
2. Detailed Comparison — field-level values, result ID, status, matching method and diagnostics.
3. Missing and Modified Items — all non-PASS rows, including extra/ambiguous/unsupported objects.
4. Drawing Images — current active graphics image and optional user-supplied reference PNG/JPEG. No all-sheet image capture or registered difference markers are claimed.
5. Validation Log — limitations, extraction failures and run messages.
6. Current Inventory — structured current records and raw API text.
7. Reference Inventory — structured reference records and raw API text.

Counts are **field/rule findings**, not counts of unique dimensions. A dimension can have a nominal PASS and tolerance FAIL. Snapshot rows do not update when drawings change. CSV text beginning with spreadsheet formula characters is prefixed with an apostrophe for safe viewing; Excel cells are explicitly text. Excel-length limits are marked; the full CSV is retained.

A unique output folder is created for each completed run. It also contains the effective company profile, optional vocabulary and log. Repeated Excel export creates a new workbook filename instead of overwriting an existing workbook.

## Establishing correspondence

A shared **string user attribute `CHECK_ID`**, unique within each category in each drawing, is the strongest intended match key. Example: the same axle bearing bore dimension in both revisions can have `CHECK_ID=BEARING_BORE_01`. Assigning this attribute is a separate manual/template workflow; this journal never writes attributes.

The next matching stages are:

- Journal identifier equality, only when you explicitly enable trust for known related revisions.
- Unique NX object name within the same type/category, with a correspondence WARNING.
- Unique exact extracted content, with a correspondence WARNING.

Every stage builds the full candidate set and accepts only one-to-one unique matches. Repeated equal dimensions remain ambiguous. Position alone never proves engineering identity. If a dimension is modified and has no stable ID/name, the safe outcome may be an ambiguous WARNING instead of a guessed comparison. A MISSING/EXTRA result is an absence within the extracted selected scope, subject to coverage errors.

Finding IDs use SHA-256 of category, rule and object identities. With CHECK_ID or a preserved journal identifier they remain stable when the compared value changes. Recreated objects/new journal identifiers can legitimately change finding IDs. Session tags are the explicitly labelled last fallback.

## Company profiles and spelling

Copy `company_profile.example.ini`, replace the placeholder string attribute titles, and select your edited file in the dialog. These are **part-attribute checks**, not verification that the displayed title block shows those values. No Dana/Teamcenter naming convention is assumed.

`Attribute.PartNumber=...` maps a string attribute. `Expected.PartNumber=...` adds an exact configured check. `AllowedSheetSizesMm=...` defines company-approved sizes explicitly. `ExpectedProjection=...` uses the exact enum value from your extraction inventory. Unknown or duplicate keys stop the run. No ISO or ASME certification is inferred from a sheet-size list.

Spelling is optional. Supply an approved UTF-8 word list, one token per line. Unknown tokens produce review warnings; engineering abbreviations, names and NX formatting tokens can be valid. The example vocabulary is deliberately incomplete. No online service or automatic wording correction is used.

## Code structure

The single journal file avoids multi-file journal compilation dependencies. Its sections are ready to separate into a compiled project later:

| Class/module | Responsibility |
|---|---|
| DrawingChecker | NX entry point, logging and transient highlight cleanup |
| Item / Finding / Settings / Snapshot | Structured data and result identity |
| Inspector | Pure matching, numeric/text field comparison and overall status |
| Extractor | Native NX traversal, typed readers, FCF builder cleanup and coverage diagnostics |
| Profiles | Explicit company rules and vocabulary review |
| CheckerForm | Setup, independent scope, progress/cancel and review grid |
| Reports / ComScope | CSV, Excel workbook generation and COM cleanup |
| EngineTests | Executable pure comparison regression tests |

## Dependencies and optional local checks

Use Windows NX 2306 with the needed Drafting/GD&T API licence features. WinForms, System.Drawing and .NET Framework are used by the journal. Excel automation is late-bound: no Excel interop DLL reference or Python package is required. Close the private Excel instance if Windows displays a first-run/activation prompt; do not use an unactivated unattended Excel setup.

The `tests` folder includes optional **Windows PowerShell 5.1** helpers:

```powershell
powershell.exe -File .\tests\Run-EngineTests.ps1
powershell.exe -File .\tests\Compile-With-NX2306.ps1 -NxManagedDirectory 'C:\Program Files\Siemens\NX2306\NXBIN\managed'
```

Find the actual folder containing your NX 2306 `NXOpen.dll` and `NXOpen.UF.dll`; the example installation path may differ. Run the engine helper in a fresh Windows PowerShell process. The compiler helper compiles but never executes the journal. Both scripts still require a local run; they were not executed here.

Use VALIDATION.md before widening the pilot scope. `sample/Sample_Inspection.xlsx` is synthetic demonstration data, not a report from your drawings or evidence of NX runtime success.
