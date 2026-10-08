' NX2306_Assembly_Exterior_Export_UI.vb
' Revision 2: STEP/Parasolid format and surface/solid/both export filters.
' Target: Siemens NX 2306, Windows x64, interactive VB.NET journal.
' Intended environment: an already-open Teamcenter-managed assembly.
'
' RUN
' 1. Fully load the precise assembly and make its top-level part Work + Display.
'    The existing assembly must be editable under your Teamcenter rules.
' 2. Tools > Journal > Play: select this file.
' 3. Run the default preflight first; inspect the Listing Window source list.
' 4. Choose STEP AP214 (.stp) or Parasolid text (.x_t), then export body type.
' 5. Clear Preflight only, choose a NEW local filename, Run.
'    If STEP settings are not found, browse to your NX AP214 ugstep214.def.
'
' RESULT
' One Linked Exterior feature set named CUSTOMER_EXTERIOR is created in the
' existing assembly work part. NX may own several sheet bodies below it.
' Only bodies owned by the exterior and matching the selected body filter export.
' Surfaces = sheet bodies; Solids = existing solid bodies; Both = both types.
' This is an EXPORT FILTER, not a sew/solidify operation. Linked Exterior normally
' produces sheets. If Solids finds none, no file is written and the feature stays.
' Re-run with Surfaces/Both to export the retained feature. Parasolid uses the
' current installed NX Parasolid version; STEP uses AP214.
' Associativity is left at the UF Linked Exterior API default, as requested.
' The component structure is retained. The journal does not create a part/item,
' check out, save, check in, or attach the STEP as a Teamcenter dataset.
' Save the modified NX assembly yourself if you want to retain its feature.
'
' SCOPE / LIMITS
' Includes root-part solids plus component solids in current reference sets,
' traversing nested/repeated occurrences in the loaded assembly arrangement.
' Suppressed and Empty-reference-set branches are skipped and counted.
' Named reference sets restrict both local bodies and child component instances.
' Hidden solid bodies are included: hiding is not an exclusion rule.
' Sheet bodies, curves, PMI, drawings and datums are not source geometry.
' Unloaded/non-precise required components and convergent solids cause a stop.
' Mixed-unit components and deformable part definitions are rejected explicitly.
' Specialized occurrence overrides/assembly cuts need a separate validation case.
' Use the assembled state, not an exploded view, and a history-mode work part.
'
' Exterior identification uses UF hidden-line visibility from 26 normalized
' directions, FINE sampling and 0.1 mm chordal tolerance by default. It copies
' whole identified faces; this is NOT a sealed wrap, cavity fill or guaranteed
' IP-removal algorithm. Faces through openings may remain; small/recessed faces
' may be missed. Inspect the generated sheets and reimported STEP before sharing.
' Openings are retained (delete_openings=False); no hole-filling is performed.
'
' FAILURE / RE-RUN
' A source-scan error aborts the whole extraction, avoiding partial assemblies.
' Extraction failure undoes this run. Export failure keeps the exterior for review
' and reports the temporary translator folder. Existing files are never replaced.
' Re-running automatically exports the existing CUSTOMER_EXTERIOR feature
' without rebuilding. To regenerate from changed assembly geometry, deliberately
' Undo/remove the old feature first. Inspect/update existing geometry in NX before
' re-exporting; this journal does not force updates of an existing feature.
' Undo does not delete an exported neutral file.
'
' VALIDATION STATUS
' API signatures checked against published NXOpen .NET reference (NX2022_1)
' and Siemens UF reference headers. NX 2306 compilation/runtime are NOT tested
' here. Native UF identification uses libufun.dll via P/Invoke because it is not
' listed in the consulted managed UFModl reference. No separate DLL is supplied.
' Requires NX solid-modeling/assembly and STEP translator capability at your site.
'
' FIRST-RUN TEST MATRIX (use non-production data)
' - Two touching blocks + separated block: one feature set, multiple sheets OK.
' - Two rotated occurrences of one part, including a nested subassembly:
'   compare positions and dimensions in the STEP against assembly ABS.
' - Closed housing with an internal solid; repeat with an opening: inspect faces.
' - Named/Empty reference sets, suppressed and hidden components: inspect log.
' - Unloaded part/read-only work part: stop before any feature or export.
' - Cancel dialogs/preflight: no model changes. Extraction error: Undo rollback.
' - STEP and Parasolid: surface-only, solid-only, both; verify imported types.
' - Solid-only with sheets: zero-match message, no empty/misleading export file.
' - Export failure: exterior retained; re-run exports it without duplication.
'
' Reference URLs:
' https://nxopencsdocumentation.thescriptingengineer.com/NX2022_1/NXOpen.UF.UFModl.html
' https://nxopencsdocumentation.thescriptingengineer.com/NX2022_1/NXOpen.UF.UFModl.LinkedExt.html
' https://www.ugapi.com/doc/ufun/uf_linked_exterior/uf_linked_exterior.html
' https://www.ugapi.com/doc/ufun/uf_modl/uf_modl_types.html
' https://www.ugapi.com/doc/ufun/uf_so/global.html
' https://nxopencsdocumentation.thescriptingengineer.com/NX2022_1/NXOpen.StepCreator.html
' https://nxopencsdocumentation.thescriptingengineer.com/NX2022_1/NXOpen.ParasolidExporter.html

Option Strict On
Option Explicit On

Imports System
Imports System.Collections.Generic
Imports System.Globalization
Imports System.IO
Imports System.Runtime.InteropServices
Imports System.Text
Imports System.Windows.Forms
Imports NXOpen
Imports NXOpen.Assemblies
Imports NXOpen.Features
Imports NXOpen.UF
Imports NXOpen.Utilities

Module AssemblyExteriorStep
    Private Enum ExportFormat
        StepAp214
        ParasolidText
    End Enum

    Private Enum BodyFilter
        SurfaceBodies
        SolidBodies
        Both
    End Enum

    Private Const ExteriorName As String = "CUSTOMER_EXTERIOR"
    Private Const HlFine As Integer = 2
    Private Const GroupSingle As Integer = 1
    Private theSession As Session
    Private uf As UFSession
    Private work As Part
    Private lw As ListingWindow
    Private sources As New List(Of SourceBody)()
    Private sourceKeys As New HashSet(Of String)()
    Private scanErrors As Integer
    Private skippedBranches As Integer
    Private skippedSheets As Integer

    Private Class SourceBody
        Public PrototypeBody As Body
        Public Occurrence As Component ' Nothing means an existing root-part body.
        Public SourcePath As String
    End Class

    ' tag_t is a 32-bit unsigned identifier in the targeted NX Windows API.
    ' Use UInt32 arrays at the native boundary and 64-bit IntPtr for UF buffers.
    <DllImport("libufun.dll", CallingConvention:=CallingConvention.Cdecl,
        EntryPoint:="UF_MODL_identify_exterior_using_hl", ExactSpelling:=True)>
    Private Function IdentifyExteriorNative(
        ByVal numberOfBodies As Integer,
        <[In]()> ByVal bodies() As UInteger,
        <[In]()> ByVal xforms() As UInteger,
        ByVal numberOfDirections As Integer,
        <[In]()> ByVal directions() As Double,
        ByVal chordalTolerance As Double,
        ByVal resolution As Integer,
        ByRef numberOfFaces As Integer,
        ByRef faces As IntPtr,
        ByRef bodyIndices As IntPtr) As Integer
    End Function

    Public Sub Main()
        theSession = Session.GetSession()
        uf = UFSession.GetUFSession()
        lw = theSession.ListingWindow
        lw.Open()
        sources.Clear()
        sourceKeys.Clear()
        scanErrors = 0
        skippedBranches = 0
        skippedSheets = 0
        Dim haveMark As Boolean = False
        Dim featureCompleted As Boolean = False
        Dim mark As Session.UndoMarkId

        Try
            Log("START " & DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
            Dim nxRelease As String = ""
            uf.UF.GetRelease(nxRelease)
            Dim managedMode As Boolean = False
            uf.UF.IsUgmanagerActive(managedMode)
            Log("NX release: " & nxRelease & "; Teamcenter-managed: " & managedMode.ToString())
            work = theSession.Parts.Work
            If work Is Nothing OrElse theSession.Parts.Display Is Nothing Then
                Throw New InvalidOperationException("Open the assembly first.")
            End If
            If work.Tag <> theSession.Parts.Display.Tag Then
                Throw New InvalidOperationException("Make the top-level assembly both Work Part and Display Part.")
            End If
            If work.IsReadOnly Then Throw New InvalidOperationException("The assembly work part is read-only. Make it editable through your normal Teamcenter process.")
            If Not work.IsFullyLoaded Then Throw New InvalidOperationException("Fully load the assembly work part before running.")
            Dim root As Component = work.ComponentAssembly.RootComponent
            If root Is Nothing OrElse root.GetChildren().Length = 0 Then
                Throw New InvalidOperationException("The current work part must contain an assembly.")
            End If
            Dim existingExterior As Tag = Tag.Null
            For Each f As Feature In work.Features
                If String.Equals(f.Name, ExteriorName, StringComparison.OrdinalIgnoreCase) Then
                    If existingExterior <> Tag.Null Then Throw New InvalidOperationException("Multiple features named " & ExteriorName & "; resolve the duplicate names before export.")
                    existingExterior = f.Tag
                End If
            Next
            Log("Work part: " & work.Leaf)
            Log("Coordinates: work/display assembly absolute; assembled positions.")
            Log("Scope: current component reference sets; nested and repeated occurrences; hidden solids included.")

            Dim preview As Boolean
            Dim toleranceMm As Double
            Dim outputPath As String
            Dim format As ExportFormat
            Dim filter As BodyFilter
            If Not GetOptions(preview, toleranceMm, outputPath, format, filter, existingExterior <> Tag.Null) Then
                Log("CANCELLED: no model changes.")
                Return
            End If

            Log("Export format: " & format.ToString() & "; body filter: " & filter.ToString())
            If existingExterior <> Tag.Null Then
                featureCompleted = True
                Log("Reusing existing " & ExteriorName & "; no rebuild or forced update.")
                Dim existingBodies As List(Of Body) = GetExteriorBodies(existingExterior, Nothing)
                Dim filteredExisting As List(Of Body) = FilterBodies(existingBodies, filter)
                If preview Then
                    Log("PREFLIGHT COMPLETE: existing exterior inspected; no file exported.")
                    MessageBox.Show("Existing exterior: " & existingBodies.Count.ToString() & " bodies." & vbCrLf &
                        "Matching export filter: " & filteredExisting.Count.ToString() & "." & vbCrLf &
                        "No model changes or export.", "Exterior preflight", MessageBoxButtons.OK, MessageBoxIcon.Information)
                    Return
                End If
                outputPath = ValidateOutputPath(outputPath, format)
                Dim existingTolerance As Double = toleranceMm
                If work.PartUnits = BasePart.Units.Inches Then existingTolerance /= 25.4
                Dim existingSettings As String = ""
                If format = ExportFormat.StepAp214 Then
                    existingSettings = GetStepSettings()
                    If existingSettings.Length = 0 Then Return
                End If
                ExportChosen(filteredExisting, outputPath, existingSettings, existingTolerance, format, filter)
                ShowExportComplete(filteredExisting.Count, outputPath)
                Return
            End If

            CollectPartBodies(work, Nothing, Nothing, "ROOT")
            For Each child As Component In root.GetChildren()
                ScanComponent(child, "ROOT/" & child.DisplayName)
            Next
            Log(String.Format(CultureInfo.InvariantCulture,
                "SOURCE SUMMARY: {0} solid occurrences; {1} skipped branches; {2} sheet bodies excluded; {3} errors.",
                sources.Count, skippedBranches, skippedSheets, scanErrors))
            If scanErrors > 0 Then Throw New InvalidOperationException("Source scan failed. Resolve the errors in the Listing Window and run again. No exterior was created.")
            If sources.Count = 0 Then Throw New InvalidOperationException("No eligible solid bodies in the current assembly/reference sets.")
            If preview Then
                Log("PREFLIGHT COMPLETE: no feature, transforms or export file created. Clear Preflight only on the next run to apply.")
                MessageBox.Show("Preflight complete: " & sources.Count.ToString() & " solid occurrences." & vbCrLf &
                    "Review the Listing Window. Run again with Preflight only cleared to create/export.",
                    "Assembly exterior", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Return
            End If

            outputPath = ValidateOutputPath(outputPath, format)
            Dim settingsFile As String = ""
            If format = ExportFormat.StepAp214 Then
                settingsFile = GetStepSettings()
                If settingsFile.Length = 0 Then
                    Log("CANCELLED at STEP settings selection: no model changes.")
                    Return
                End If
            End If
            Dim tolerance As Double = toleranceMm
            If work.PartUnits = BasePart.Units.Inches Then tolerance /= 25.4
            Dim before As New HashSet(Of Tag)()
            For Each b As Body In work.Bodies
                before.Add(b.Tag)
            Next
            mark = theSession.SetUndoMark(Session.MarkVisibility.Visible, "Create customer linked exterior")
            haveMark = True

            Dim bodyTags(sources.Count - 1) As Tag
            Dim xformTags(sources.Count - 1) As Tag
            Dim transforms As New Dictionary(Of Tag, Tag)()
            For i As Integer = 0 To sources.Count - 1
                bodyTags(i) = sources(i).PrototypeBody.Tag
                If sources(i).Occurrence IsNot Nothing Then
                    Dim occurrenceTag As Tag = sources(i).Occurrence.Tag
                    Dim transformTag As Tag = Tag.Null
                    If Not transforms.TryGetValue(occurrenceTag, transformTag) Then
                        uf.So.CreateXformAssyCtxt(work.Tag, occurrenceTag, Tag.Null, transformTag)
                        transforms.Add(occurrenceTag, transformTag)
                    End If
                    xformTags(i) = transformTag
                Else
                    xformTags(i) = Tag.Null
                End If
            Next

            Log("Identifying exterior: 26 directions, FINE, chordal tolerance " & toleranceMm.ToString(CultureInfo.InvariantCulture) & " mm.")
            Dim faces() As Tag = Nothing
            Dim indices() As Integer = Nothing
            IdentifyExterior(bodyTags, xformTags, tolerance, faces, indices)
            Log("Identified face occurrences: " & faces.Length.ToString())
            Dim data As New UFModl.LinkedExt()
            data.num_bodies = bodyTags.Length
            data.bodies = bodyTags
            data.xforms = xformTags
            data.num_faces = faces.Length
            data.faces = faces
            data.xform_index = indices
            data.group_results = GroupSingle
            data.mass_props = False
            data.delete_openings = False
            data.at_timestamp = False ' Update timing only; NOT an associativity switch.
            Dim featureTag As Tag = Tag.Null
            uf.Modl.CreateLinkedExterior(data, featureTag)
            If featureTag = Tag.Null Then Throw New InvalidOperationException("NX returned no Linked Exterior feature.")
            Dim exterior As NXObject = TryCast(NXObjectManager.Get(featureTag), NXObject)
            If exterior Is Nothing Then Throw New InvalidOperationException("Unable to obtain the created Linked Exterior object.")
            exterior.SetName(ExteriorName)
            Dim updateErrors As Integer = theSession.UpdateManager.DoUpdate(mark)
            If updateErrors <> 0 Then Throw New InvalidOperationException("NX update reported " & updateErrors.ToString() & " error(s).")

            Dim resultBodies As List(Of Body) = GetExteriorBodies(featureTag, before)
            For i As Integer = 0 To resultBodies.Count - 1
                resultBodies(i).SetName("CUSTOMER_SURFACE_" & (i + 1).ToString("0000", CultureInfo.InvariantCulture))
            Next
            featureCompleted = True
            Log("CREATED " & ExteriorName & ": " & resultBodies.Count.ToString() & " bodies. Associativity left at API default.")
            Log("The feature is in the existing work part; the assembly has not been saved.")
            Log("ENGINEERING REVIEW: check openings, internal faces, small details and assembly position before customer release.")

            Dim exportBodies As List(Of Body) = FilterBodies(resultBodies, filter)
            ExportChosen(exportBodies, outputPath, settingsFile, tolerance, format, filter)
            ShowExportComplete(exportBodies.Count, outputPath)
        Catch ex As Exception
            Log("ERROR: " & ex.GetType().Name & ": " & ex.Message)
            If haveMark AndAlso Not featureCompleted Then
                Try
                    theSession.UndoToMark(mark, Nothing)
                    Log("ROLLED BACK this run's extraction changes.")
                Catch rollbackEx As Exception
                    Log("ROLLBACK ERROR: " & rollbackEx.Message & ". Inspect the assembly before saving.")
                End Try
            ElseIf featureCompleted Then
                Log("The exterior is retained. Export was NOT confirmed. Re-run to export this feature with revised options; no duplicate will be created.")
            End If
            MessageBox.Show(ex.Message & vbCrLf & vbCrLf & "See the NX Listing Window for details.",
                "Assembly exterior stopped", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub ScanComponent(ByVal c As Component, ByVal path As String)
        Try
            If c.IsSuppressed Then
                skippedBranches += 1
                Log("SKIP suppressed branch: " & path)
                Return
            End If
            If String.Equals(c.ReferenceSet, "Empty", StringComparison.OrdinalIgnoreCase) Then
                skippedBranches += 1
                Log("SKIP Empty reference set: " & path)
                Return
            End If
            Dim p As Part = TryCast(c.Prototype, Part)
            If p Is Nothing OrElse Not p.IsFullyLoaded Then
                Throw New InvalidOperationException("Required component is not a fully loaded precise NX part.")
            End If
            If p.PartUnits <> work.PartUnits Then Throw New InvalidOperationException("Mixed-unit assembly is not supported by this journal version.")
            If uf.Assem.IsPartDeformable(p.Tag) Then Throw New InvalidOperationException("Deformable component needs a separately validated extraction workflow.")
            Dim members As HashSet(Of Tag) = ReferenceMembers(p, c.ReferenceSet)
            CollectPartBodies(p, c, members, path)
            For Each child As Component In c.GetChildren()
                Dim includeChild As Boolean = (members Is Nothing)
                If Not includeChild Then
                    Dim instanceTag As Tag = uf.Assem.AskInstOfPartOcc(child.Tag)
                    includeChild = members.Contains(instanceTag)
                End If
                If includeChild Then
                    ScanComponent(child, path & "/" & child.DisplayName)
                Else
                    skippedBranches += 1
                    Log("SKIP outside parent reference set: " & path & "/" & child.DisplayName)
                End If
            Next
        Catch ex As Exception
            scanErrors += 1
            Log("SOURCE ERROR " & path & ": " & ex.Message)
        End Try
    End Sub

    Private Function ReferenceMembers(ByVal p As Part, ByVal refName As String) As HashSet(Of Tag)
        If String.Equals(refName, "Entire Part", StringComparison.OrdinalIgnoreCase) Then Return Nothing
        For Each rs As ReferenceSet In p.GetAllReferenceSets()
            If String.Equals(rs.Name, refName, StringComparison.Ordinal) Then
                Dim result As New HashSet(Of Tag)()
                For Each member As NXObject In rs.AskMembersInReferenceSet()
                    result.Add(member.Tag)
                Next
                Return result
            End If
        Next
        Throw New InvalidOperationException("Cannot resolve reference set '" & refName & "'. No fallback to Entire Part is used.")
    End Function

    Private Sub CollectPartBodies(ByVal p As Part, ByVal c As Component,
        ByVal members As HashSet(Of Tag), ByVal path As String)
        For Each b As Body In p.Bodies
            Try
                If b.IsOccurrence Then Continue For
                If members IsNot Nothing AndAlso Not members.Contains(b.Tag) Then Continue For
                If b.IsConvergentBody Then Throw New InvalidOperationException("Convergent body is not supported: " & b.Tag.ToString())
                If Not b.IsSolidBody Then
                    skippedSheets += 1
                    Continue For
                End If
                If b.GetFaces().Length = 0 Then Throw New InvalidOperationException("Solid has no accessible precise faces: " & b.Tag.ToString())
                Dim ownerKey As String = If(c Is Nothing, "ROOT", c.Tag.ToString())
                Dim key As String = ownerKey & ":" & b.Tag.ToString()
                If sourceKeys.Add(key) Then
                    Dim item As New SourceBody()
                    item.PrototypeBody = b
                    item.Occurrence = c
                    item.SourcePath = path
                    sources.Add(item)
                    Log("SOURCE " & path & " | body " & b.Tag.ToString())
                End If
            Catch ex As Exception
                scanErrors += 1
                Log("BODY ERROR " & path & ": " & ex.Message)
            End Try
        Next
    End Sub

    Private Sub IdentifyExterior(ByVal bodies() As Tag, ByVal xforms() As Tag,
        ByVal tolerance As Double, ByRef faceTags() As Tag, ByRef indices() As Integer)
        If Marshal.SizeOf(System.Enum.GetUnderlyingType(GetType(Tag))) <> 4 Then
            Throw New NotSupportedException("Unexpected NX tag width; native identification was not called.")
        End If
        Dim rawBodies(bodies.Length - 1) As UInteger
        Dim rawXforms(xforms.Length - 1) As UInteger
        For i As Integer = 0 To bodies.Length - 1
            rawBodies(i) = Convert.ToUInt32(bodies(i), CultureInfo.InvariantCulture)
            rawXforms(i) = Convert.ToUInt32(xforms(i), CultureInfo.InvariantCulture)
        Next
        Dim dirs As New List(Of Double)()
        For x As Integer = -1 To 1
            For y As Integer = -1 To 1
                For z As Integer = -1 To 1
                    If x = 0 AndAlso y = 0 AndAlso z = 0 Then Continue For
                    Dim length As Double = Math.Sqrt(CDbl(x * x + y * y + z * z))
                    dirs.Add(x / length)
                    dirs.Add(y / length)
                    dirs.Add(z / length)
                Next
            Next
        Next
        Dim count As Integer = 0
        Dim faceMemory As IntPtr = IntPtr.Zero
        Dim indexMemory As IntPtr = IntPtr.Zero
        Try
            Dim rc As Integer = IdentifyExteriorNative(bodies.Length, rawBodies, rawXforms,
                dirs.Count \ 3, dirs.ToArray(), tolerance, HlFine, count, faceMemory, indexMemory)
            If rc <> 0 Then
                Dim detail As String = ""
                uf.UF.GetFailMessage(rc, detail)
                Throw New InvalidOperationException("Exterior identification error " & rc.ToString() & ": " & detail)
            End If
            If count <= 0 OrElse faceMemory = IntPtr.Zero OrElse indexMemory = IntPtr.Zero Then
                Throw New InvalidOperationException("NX identified no exterior faces.")
            End If
            ReDim faceTags(count - 1)
            ReDim indices(count - 1)
            Marshal.Copy(indexMemory, indices, 0, count)
            For i As Integer = 0 To count - 1
                Dim rawSigned As Integer = Marshal.ReadInt32(faceMemory, i * 4)
                Dim rawUnsigned As UInteger = BitConverter.ToUInt32(BitConverter.GetBytes(rawSigned), 0)
                faceTags(i) = CType(System.Enum.ToObject(GetType(Tag), rawUnsigned), Tag)
                If indices(i) < 0 OrElse indices(i) >= bodies.Length Then
                    Throw New InvalidOperationException("Invalid exterior face-to-body index returned by NX.")
                End If
                Dim f As Face = TryCast(NXObjectManager.Get(faceTags(i)), Face)
                If f Is Nothing OrElse f.IsOccurrence Then Throw New InvalidOperationException("Expected a prototype face from the exterior API.")
                If f.GetBody().Tag <> bodies(indices(i)) Then
                    Throw New InvalidOperationException("Exterior face does not match its reported prototype body.")
                End If
            Next
        Finally
            If faceMemory <> IntPtr.Zero Then uf.UF.Free(faceMemory)
            If indexMemory <> IntPtr.Zero Then uf.UF.Free(indexMemory)
        End Try
    End Sub

    Private Function GetExteriorBodies(ByVal exteriorTag As Tag,
        ByVal before As HashSet(Of Tag)) As List(Of Body)
        Dim extData As New UFModl.LinkedExt()
        Dim groupCount As Integer = 0
        Dim groups() As Tag = Nothing
        Dim subCount As Integer = 0
        Dim subFeatures() As Tag = Nothing
        Dim mass(46) As Double
        uf.Modl.AskLinkedExterior(exteriorTag, extData, groupCount, groups, subCount, subFeatures, mass)
        If subCount <= 0 OrElse subFeatures Is Nothing Then Throw New InvalidOperationException("Linked Exterior returned no owned subfeatures.")
        Dim bodiesByTag As New Dictionary(Of Tag, Body)()
        For Each subTag As Tag In subFeatures
            Dim f As Feature = TryCast(NXObjectManager.Get(subTag), Feature)
            If f Is Nothing Then Throw New InvalidOperationException("Cannot inspect a Linked Exterior subfeature.")
            For Each b As Body In f.GetBodies()
                If b.IsOccurrence OrElse b.OwningPart.Tag <> work.Tag OrElse (before IsNot Nothing AndAlso before.Contains(b.Tag)) Then
                    Throw New InvalidOperationException("Exterior export validation found a source or non-work-part body.")
                End If
                If Not bodiesByTag.ContainsKey(b.Tag) Then bodiesByTag.Add(b.Tag, b)
            Next
        Next
        If bodiesByTag.Count = 0 Then Throw New InvalidOperationException("No owned exterior bodies to export.")
        If before IsNot Nothing Then
            For Each b As Body In work.Bodies
                If Not before.Contains(b.Tag) AndAlso Not bodiesByTag.ContainsKey(b.Tag) Then
                    Throw New InvalidOperationException("Unexpected additional body appeared during update. Export stopped for review.")
                End If
            Next
        End If
        Return New List(Of Body)(bodiesByTag.Values)
    End Function

    Private Function FilterBodies(ByVal bodies As List(Of Body), ByVal filter As BodyFilter) As List(Of Body)
        Dim result As New List(Of Body)()
        Dim solids As Integer = 0
        Dim sheets As Integer = 0
        For Each b As Body In bodies
            If b.IsSolidBody Then
                solids += 1
                If filter <> BodyFilter.SurfaceBodies Then result.Add(b)
            ElseIf b.IsSheetBody Then
                sheets += 1
                If filter <> BodyFilter.SolidBodies Then result.Add(b)
            Else
                Throw New InvalidOperationException("An exterior body has an unsupported body type: " & b.Tag.ToString())
            End If
        Next
        Log("EXTERIOR TYPES: " & solids.ToString() & " solid bodies; " & sheets.ToString() &
            " surface bodies; " & result.Count.ToString() & " match " & filter.ToString() & ".")
        If result.Count = 0 Then
            Throw New InvalidOperationException("No exterior bodies match " & filter.ToString() & "." & vbCrLf &
                "The exterior contains " & solids.ToString() & " solids and " & sheets.ToString() & " surface bodies." & vbCrLf &
                "This option filters bodies; it does not convert surfaces into solids. No file was exported.")
        End If
        Return result
    End Function

    Private Sub ExportChosen(ByVal bodies As List(Of Body), ByVal outputPath As String,
        ByVal settingsFile As String, ByVal tolerance As Double, ByVal format As ExportFormat,
        ByVal filter As BodyFilter)
        If bodies.Count = 0 Then Throw New InvalidOperationException("No bodies selected for export.")
        If format = ExportFormat.StepAp214 Then
            ExportStep(bodies, outputPath, settingsFile, tolerance, filter)
        Else
            ExportParasolid(bodies, outputPath, filter)
        End If
        Log("EXPORT COMPLETE: " & outputPath)
        Log("Only the filtered exterior bodies were selected. No assembly save/check-in performed.")
    End Sub

    Private Sub ShowExportComplete(ByVal count As Integer, ByVal outputPath As String)
        MessageBox.Show("Exported " & count.ToString() & " bodies from " & ExteriorName & "." & vbCrLf &
            outputPath & vbCrLf & vbCrLf &
            "Inspect the exported model before sharing. Save the NX assembly manually if needed.",
            "Exterior export complete", MessageBoxButtons.OK, MessageBoxIcon.Information)
    End Sub

    Private Sub ExportParasolid(ByVal bodies As List(Of Body), ByVal outputPath As String,
        ByVal filter As BodyFilter)
        Dim tempFolder As String = Path.Combine(Path.GetDirectoryName(outputPath),
            "NXExterior_" & Guid.NewGuid().ToString("N"))
        Directory.CreateDirectory(tempFolder)
        Dim tempFile As String = Path.Combine(tempFolder, "customer_exterior.x_t")
        Dim exporter As ParasolidExporter = Nothing
        Dim success As Boolean = False
        Try
            exporter = theSession.DexManager.CreateParasolidExporter()
            exporter.ExportDestination = BaseCreator.ExportDestinationOption.NativeFileSystem
            exporter.ExportFrom = ParasolidExporter.ExportFromOption.DisplayedPart
            exporter.InputFile = ""
            exporter.OutputFile = tempFile
            exporter.ProcessHoldFlag = True
            exporter.FlattenAssembly = True
            exporter.ParasolidVersion = ParasolidExporter.ParasolidVersionOption.Current
            exporter.ObjectTypes.Solids = (filter <> BodyFilter.SurfaceBodies)
            exporter.ObjectTypes.Surfaces = (filter <> BodyFilter.SolidBodies)
            exporter.ObjectTypes.Curves = False
            exporter.ObjectTypes.Annotations = False
            exporter.ObjectTypes.Csys = False
            exporter.ObjectTypes.PmiData = False
            exporter.ObjectTypes.Attributes = False
            exporter.ObjectTypes.ProductData = False
            exporter.ObjectTypes.Structures = False
            exporter.ObjectTypes.Kinematic = False
            exporter.ObjectTypes.FacetBodies = False
            exporter.ExportSelectionBlock.SelectionScope = ObjectSelector.Scope.SelectedObjects
            Dim selected(bodies.Count - 1) As NXObject
            For i As Integer = 0 To bodies.Count - 1
                selected(i) = bodies(i)
            Next
            If Not exporter.ExportSelectionBlock.SelectionComp.Add(selected) Then
                Throw New InvalidOperationException("Parasolid exporter rejected the exterior body selection.")
            End If
            Log("Parasolid text export: " & selected.Length.ToString() & " selected exterior bodies; current NX Parasolid version.")
            exporter.Commit()
            If Not File.Exists(tempFile) Then Throw New IOException("Parasolid exporter did not create the requested file.")
            Dim resultInfo As New FileInfo(tempFile)
            If resultInfo.Length = 0 Then Throw New IOException("Parasolid exporter produced an empty file.")
            ' File existence/size is checked here, not kernel validity. Reimport in NX to validate.
            File.Move(tempFile, outputPath)
            success = True
        Catch
            Log("Parasolid working folder retained for diagnosis: " & tempFolder)
            Throw
        Finally
            If exporter IsNot Nothing Then
                Try
                    exporter.Destroy()
                Catch cleanupEx As Exception
                    Log("Parasolid builder cleanup warning: " & cleanupEx.Message)
                End Try
            End If
            If success Then
                Try
                    Directory.Delete(tempFolder, True)
                Catch cleanupEx As Exception
                    Log("NOTE: Parasolid temporary files remain at " & tempFolder & ": " & cleanupEx.Message)
                End Try
            End If
        End Try
    End Sub

    Private Sub ExportStep(ByVal bodies As List(Of Body), ByVal outputPath As String,
        ByVal settingsFile As String, ByVal tolerance As Double, ByVal filter As BodyFilter)
        Dim tempFolder As String = Path.Combine(Path.GetDirectoryName(outputPath),
            "NXExterior_" & Guid.NewGuid().ToString("N"))
        Directory.CreateDirectory(tempFolder)
        Dim tempStep As String = Path.Combine(tempFolder, "customer_exterior.stp")
        Dim creator As StepCreator = Nothing
        Dim success As Boolean = False
        Try
            creator = theSession.DexManager.CreateStepCreator()
            creator.SettingsFile = settingsFile
            creator.ExportAs = StepCreator.ExportAsOption.Ap214
            creator.ExportDestination = BaseCreator.ExportDestinationOption.NativeFileSystem
            creator.ExportFrom = StepCreator.ExportFromOption.DisplayPart
            creator.InputFile = "" ' Use current in-memory display part, not a saved @DB path.
            creator.OutputFile = tempStep
            creator.FileSaveFlag = False
            creator.ProcessHoldFlag = True
            creator.ExportExtRef = False
            creator.LayerMask = "1-256"
            creator.BsplineTol = tolerance
            creator.ObjectTypes.Solids = (filter <> BodyFilter.SurfaceBodies)
            creator.ObjectTypes.Surfaces = (filter <> BodyFilter.SolidBodies)
            creator.ObjectTypes.Curves = False
            creator.ObjectTypes.Annotations = False
            creator.ObjectTypes.Csys = False
            creator.ObjectTypes.PmiData = False
            creator.ObjectTypes.Attributes = False
            creator.ObjectTypes.ProductData = False
            creator.ObjectTypes.Structures = False
            creator.ObjectTypes.Kinematic = False
            creator.ObjectTypes.FacetBodies = False
            creator.UserDefinedAttributes = False
            creator.SystemAttributes = False
            creator.ValidationProperties = False
            creator.AssemblyValidationProperties = False
            creator.GeometricValidationProperties = False
            creator.PMIValidationProperties = False
            creator.Author = ""
            creator.Company = ""
            creator.Authorization = ""
            creator.Description = "Customer exterior surfaces"
            creator.ExportSelectionBlock.SelectionScope = ObjectSelector.Scope.SelectedObjects
            Dim selected(bodies.Count - 1) As NXObject
            For i As Integer = 0 To bodies.Count - 1
                selected(i) = bodies(i)
            Next
            If Not creator.ExportSelectionBlock.SelectionComp.Add(selected) Then
                Throw New InvalidOperationException("STEP translator did not accept the exterior body selection.")
            End If
            Log("STEP export: " & selected.Length.ToString() & " explicitly selected exterior bodies; AP214.")
            creator.Commit()
            If Not File.Exists(tempStep) Then Throw New IOException("STEP translator did not create the requested file.")
            ValidateStepFile(tempStep)
            ' File.Move will fail if the destination exists; never overwrite.
            File.Move(tempStep, outputPath)
            success = True
        Catch
            Log("STEP translator working folder retained for diagnosis: " & tempFolder)
            Throw
        Finally
            If creator IsNot Nothing Then
                Try
                    creator.Destroy()
                Catch cleanupEx As Exception
                    Log("STEP builder cleanup warning: " & cleanupEx.Message)
                End Try
            End If
            If success Then
                Try
                    Directory.Delete(tempFolder, True)
                Catch ex As Exception
                    Log("NOTE: translator temporary files remain at " & tempFolder & ": " & ex.Message)
                End Try
            End If
        End Try
    End Sub

    Private Sub ValidateStepFile(ByVal filename As String)
        Using stream As New FileStream(filename, FileMode.Open, FileAccess.Read, FileShare.Read)
            If stream.Length < 100 Then Throw New IOException("STEP file is empty or too short.")
            Dim length As Integer = CInt(Math.Min(4096L, stream.Length))
            Dim header(length - 1) As Byte
            stream.Read(header, 0, length)
            If Not Encoding.ASCII.GetString(header).Contains("ISO-10303-21;") Then
                Throw New IOException("STEP header is missing.")
            End If
            stream.Seek(-CLng(length), SeekOrigin.End)
            Dim tail(length - 1) As Byte
            stream.Read(tail, 0, length)
            If Not Encoding.ASCII.GetString(tail).Contains("END-ISO-10303-21;") Then
                Throw New IOException("STEP end marker is missing; translation may be incomplete.")
            End If
        End Using
    End Sub

    Private Function ValidateOutputPath(ByVal rawPath As String, ByVal format As ExportFormat) As String
        If String.IsNullOrWhiteSpace(rawPath) Then Throw New ArgumentException("Choose a local export filename.")
        If Not Path.IsPathRooted(rawPath) Then Throw New ArgumentException("Use an absolute output path.")
        Dim result As String = Path.GetFullPath(rawPath)
        Dim ext As String = Path.GetExtension(result).ToLowerInvariant()
        If format = ExportFormat.StepAp214 Then
            If ext <> ".stp" AndAlso ext <> ".step" Then Throw New ArgumentException("STEP output must have a .stp or .step extension.")
        Else
            If ext <> ".x_t" Then Throw New ArgumentException("Parasolid text output must have a .x_t extension.")
        End If
        If Not Directory.Exists(Path.GetDirectoryName(result)) Then Throw New DirectoryNotFoundException("Output directory does not exist.")
        If File.Exists(result) Then Throw New IOException("Output already exists. Choose a new filename; existing files are never replaced.")
        Return result
    End Function

    Private Function GetStepSettings() As String
        Dim candidates As New List(Of String)()
        Dim stepDir As String = theSession.GetEnvironmentVariableValue("UGII_STEP214_DIR")
        If Not String.IsNullOrWhiteSpace(stepDir) Then candidates.Add(Path.Combine(stepDir, "ugstep214.def"))
        Dim baseDir As String = theSession.GetEnvironmentVariableValue("UGII_BASE_DIR")
        If Not String.IsNullOrWhiteSpace(baseDir) Then candidates.Add(Path.Combine(baseDir, "step214ug", "ugstep214.def"))
        For Each candidate As String In candidates
            If File.Exists(candidate) Then Return candidate
        Next
        Using dialog As New OpenFileDialog()
            dialog.Title = "Select the NX AP214 translator settings (ugstep214.def)"
            dialog.Filter = "STEP translator settings (*.def)|*.def"
            dialog.CheckFileExists = True
            If dialog.ShowDialog() = DialogResult.OK Then Return dialog.FileName
        End Using
        Return ""
    End Function

    Private Function GetOptions(ByRef preview As Boolean, ByRef toleranceMm As Double,
        ByRef outputPath As String, ByRef format As ExportFormat, ByRef filter As BodyFilter,
        ByVal reuseExisting As Boolean) As Boolean
        Using form As New Form()
            form.Text = "NX customer exterior - STEP / Parasolid"
            form.ClientSize = New System.Drawing.Size(660, 400)
            form.FormBorderStyle = FormBorderStyle.FixedDialog
            form.StartPosition = FormStartPosition.CenterScreen
            form.AutoScaleMode = AutoScaleMode.Dpi
            form.MaximizeBox = False
            form.MinimizeBox = False
            Dim scope As New Label()
            scope.SetBounds(18, 14, 622, 40)
            scope.Text = If(reuseExisting,
                "Existing CUSTOMER_EXTERIOR found: export its current bodies without rebuilding.",
                "Create one Linked Exterior in the existing assembly using current component reference sets.")
            Dim check As New CheckBox()
            check.SetBounds(18, 56, 620, 25)
            check.Text = "Preflight only - inspect sources/result without creating or exporting"
            check.Checked = True
            Dim formatLabel As New Label()
            formatLabel.SetBounds(18, 100, 190, 24)
            formatLabel.Text = "Export format"
            Dim formatBox As New ComboBox()
            formatBox.SetBounds(220, 96, 420, 28)
            formatBox.DropDownStyle = ComboBoxStyle.DropDownList
            formatBox.Items.AddRange(New Object() {"STEP AP214 (.stp)", "Parasolid text (.x_t) - current NX version"})
            formatBox.SelectedIndex = 0
            Dim typeLabel As New Label()
            typeLabel.SetBounds(18, 140, 190, 24)
            typeLabel.Text = "Bodies to export"
            Dim typeBox As New ComboBox()
            typeBox.SetBounds(220, 136, 420, 28)
            typeBox.DropDownStyle = ComboBoxStyle.DropDownList
            typeBox.Items.AddRange(New Object() {"Surface bodies (sheets)", "Solid bodies only", "Both solids and surfaces"})
            typeBox.SelectedIndex = 0
            Dim help As New Label()
            help.SetBounds(18, 178, 622, 43)
            help.Text = "Body type filters the exterior result. It does not sew surfaces or create solids. If no bodies match, no file is exported."
            Dim tolLabel As New Label()
            tolLabel.SetBounds(18, 227, 235, 24)
            tolLabel.Text = "Chordal / STEP tolerance (mm)"
            Dim tol As New NumericUpDown()
            tol.SetBounds(260, 224, 110, 25)
            tol.DecimalPlaces = 3
            tol.Minimum = 0.001D
            tol.Maximum = 10D
            tol.Increment = 0.01D
            tol.Value = 0.1D
            Dim pathBox As New TextBox()
            pathBox.SetBounds(18, 271, 520, 26)
            pathBox.Text = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory), "Customer_Exterior.stp")
            AddHandler formatBox.SelectedIndexChanged, Sub(sender As Object, e As EventArgs)
                Try
                    pathBox.Text = Path.ChangeExtension(pathBox.Text, If(formatBox.SelectedIndex = 0, ".stp", ".x_t"))
                Catch ex As ArgumentException
                    pathBox.Text = ""
                End Try
            End Sub
            Dim browse As New Button()
            browse.SetBounds(548, 269, 92, 29)
            browse.Text = "Browse..."
            AddHandler browse.Click, Sub(sender As Object, e As EventArgs)
                Using save As New SaveFileDialog()
                    Dim isStep As Boolean = (formatBox.SelectedIndex = 0)
                    save.Title = "Choose a NEW export filename"
                    save.Filter = If(isStep, "STEP files (*.stp)|*.stp|STEP files (*.step)|*.step", "Parasolid text (*.x_t)|*.x_t")
                    save.FileName = If(isStep, "Customer_Exterior.stp", "Customer_Exterior.x_t")
                    save.DefaultExt = If(isStep, "stp", "x_t")
                    save.AddExtension = True
                    save.OverwritePrompt = True
                    If save.ShowDialog(form) = DialogResult.OK Then pathBox.Text = save.FileName
                End Using
            End Sub
            Dim note As New Label()
            note.SetBounds(18, 311, 622, 30)
            note.Text = "Local file export. Existing files are not replaced. Save the NX assembly manually if needed."
            Dim run As New Button()
            run.SetBounds(446, 356, 92, 29)
            run.Text = "Run"
            run.DialogResult = DialogResult.OK
            Dim cancel As New Button()
            cancel.SetBounds(548, 356, 92, 29)
            cancel.Text = "Cancel"
            cancel.DialogResult = DialogResult.Cancel
            form.Controls.AddRange(New Control() {scope, check, formatLabel, formatBox, typeLabel, typeBox,
                help, tolLabel, tol, pathBox, browse, note, run, cancel})
            form.AcceptButton = run
            form.CancelButton = cancel
            If form.ShowDialog() <> DialogResult.OK Then Return False
            preview = check.Checked
            toleranceMm = CDbl(tol.Value)
            outputPath = pathBox.Text.Trim()
            format = CType(formatBox.SelectedIndex, ExportFormat)
            filter = CType(typeBox.SelectedIndex, BodyFilter)
            Return True
        End Using
    End Function

    Private Sub Log(ByVal text As String)
        lw.WriteLine(text)
    End Sub

    Public Function GetUnloadOption(ByVal dummy As String) As Integer
        Return CInt(Session.LibraryUnloadOption.Immediately)
    End Function
End Module
