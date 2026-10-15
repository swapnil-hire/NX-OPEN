# API evidence and local compatibility checks

These are mirrors of Siemens-authored NXOpen generated reference documentation. The `SC2306` paths identify the 2306 documentation set. Some legacy class pages are available only from the older `NX2022_1` mirror. They establish documented interfaces, not a successful binary/runtime test in your installation.

| API / purpose | Reference | Status |
|---|---|---|
| Dimension.ComputedSize; upper/lower tolerance; precision; fit values; GetDimensionText; GetMeasurementType; GetDimensionPreferences; GetAppendedText; first/second associativity | https://nxopencsdocumentation.thescriptingengineer.com/SC2306/NXOpen.Annotations.Dimension.html | SC2306 declaration checked; subtype/unit semantics need fixtures |
| Annotation origin, state, GetViews, NumberOfAssociativities | https://nxopencsdocumentation.thescriptingengineer.com/SC2306/NXOpen.Annotations.Annotation.html | SC2306 declarations checked |
| GetViews is PMI-specific and can return empty for ordinary drafting | https://www.ugapi.com/doc/nxopen/a00005.html | Siemens C++ semantics checked; not treated as universal sheet ownership |
| UF.View.AskViewDependentStatus | https://nxopencsdocumentation.thescriptingengineer.com/NX2022_1/NXOpen.UF.UFView.html | Older documented managed signature |
| Native view dependency status (uc6409): 0 not dependent, 1 dependent, associated view name | https://www.ugapi.com/doc/ufun/uf_view/global.html | Siemens UF semantics; name must resolve uniquely in owner-part map |
| DrawingSheet length/height, units, projection, GetScale, GetDraftingViews, View | https://nxopencsdocumentation.thescriptingengineer.com/SC2306/NXOpen.Drawings.DrawingSheet.html | SC2306 declarations checked |
| DrawingSheetCollection.ToArray | https://nxopencsdocumentation.thescriptingengineer.com/NX2022_1/NXOpen.Drawings.DrawingSheetCollection.html | Legacy collection remains a local compile checkpoint |
| DraftingView subtype, origin, broken/out-of-date flags and Style | https://nxopencsdocumentation.thescriptingengineer.com/SC2306/NXOpen.Drawings.DraftingView.html | SC2306 declarations checked |
| GeneralViewStyle.Scale/Angle/ViewLabel/ScaleLabel | https://nxopencsdocumentation.thescriptingengineer.com/NX2022_1/NXOpen.Drawings.GeneralViewStyle.html | Older interface; local compile/runtime check |
| HiddenLinesViewStyle.Hiddenline | https://nxopencsdocumentation.thescriptingengineer.com/NX2022_1/NXOpen.Drawings.HiddenLinesViewStyle.html | Exact spelling/case checked |
| AnnotationManager.CreateDraftingFeatureControlFrameBuilder(Gdt), CreatePmiFeatureControlFrameBuilder(Fcf) | https://nxopencsdocumentation.thescriptingengineer.com/SC2306/NXOpen.Annotations.AnnotationManager.html | SC2306 declarations checked; builders destroyed without Commit |
| FeatureControlFrameBuilder characteristic, frame style and FeatureControlFrameData(out array) | https://nxopencsdocumentation.thescriptingengineer.com/SC2306/NXOpen.Annotations.FeatureControlFrameBuilder.html | SC2306 declarations checked |
| FeatureControlFrameDataBuilder scalar properties, datum roles, compound references | https://nxopencsdocumentation.thescriptingengineer.com/SC2306/NXOpen.Annotations.FeatureControlFrameDataBuilder.html | SC2306 property whitelist; reflection reports each missing/read failure |
| DatumReferenceBuilder letter/material/free-state/projected properties | https://nxopencsdocumentation.thescriptingengineer.com/SC2306/NXOpen.Annotations.DatumReferenceBuilder.html | SC2306 declarations checked |
| CompoundDatumReferenceBuilder.DatumReferenceList | https://nxopencsdocumentation.thescriptingengineer.com/SC2306/NXOpen.Annotations.CompoundDatumReferenceBuilder.html | SC2306 declaration checked |
| DatumReferenceBuilderList.GetContents | https://nxopencsdocumentation.thescriptingengineer.com/NX2022_1/NXOpen.Annotations.DatumReferenceBuilderList.html | Older declaration; local compile checkpoint |
| BaseDatum.Label / DraftingDatum | https://nxopencsdocumentation.thescriptingengineer.com/NX2022_1/NXOpen.Annotations.BaseDatum.html | Older declarations; local compile checkpoint |
| SimpleDraftingAid.GetText inherited through BaseFcf | https://nxopencsdocumentation.thescriptingengineer.com/NX2022_1/NXOpen.Annotations.BaseFcf.html | Older hierarchy; local compile checkpoint |
| AppendedText.GetBeforeText/GetAfterText/GetAboveText/GetBelowText; IDisposable | https://nxopencsdocumentation.thescriptingengineer.com/SC2306/NXOpen.Annotations.AppendedText.html | SC2306 declarations checked |
| DimensionPreferences.GetUnitsFormatPreferences | https://nxopencsdocumentation.thescriptingengineer.com/SC2306/NXOpen.Annotations.DimensionPreferences.html | SC2306 declaration checked |
| UnitsFormatPreferences.PrimaryDimensionUnit, DualDimensionUnit, AngularDimensionFormat | https://nxopencsdocumentation.thescriptingengineer.com/NX2022_1/NXOpen.Annotations.UnitsFormatPreferences.html | Older interface, disposable reader |
| UF.Obj.CycleAll, AskTypeAndSubtype | https://nxopencsdocumentation.thescriptingengineer.com/NX2022_1/NXOpen.UF.UFObj.html | Older managed signatures checked |
| UF drafting entity and dimension type constants | https://nxopencsdocumentation.thescriptingengineer.com/NX12/NXOpen.UF.UFConstants.html | Legacy named constants, not guessed numeric literals |
| PartCollection.Open / PartLoadStatus disposal | https://nxopencsdocumentation.thescriptingengineer.com/NX2022_1/NXOpen.PartCollection.html | Opens without intentional work/display switching; verify with installed NX |
| ImageExportBuilder PNG, filename, RegionMode, BackgroundOption | https://nxopencsdocumentation.thescriptingengineer.com/NX2022_1/NXOpen.Gateway.ImageExportBuilder.html | Image-only Commit; external file operation |

## Required local compatibility procedure

Run `tests/Compile-With-NX2306.ps1` with the actual NX 2306 managed DLL folder. Record DLL file versions and compiler errors. This checks declarations, namespaces and references. Then run in NX with known fixtures; compilation cannot validate builder behaviour, background-part ownership or measured/displayed units.

Record a short NX journal for any subtype whose fields do not extract reliably: a simple datum, a two-segment/composite FCF, a common datum, projected zone, an overridden dimension, a hole/thread callout and your actual title block. Use the installed NXOpen help to validate the exact read adapter. Do not replace unavailable members with same-sounding guessed APIs.

Standards rules require separately verified and configured editions/clauses. This package does not quote or implement unverified ASME/ISO requirements.
