Option Strict On
Option Explicit On

Imports System
Imports System.IO
Imports System.Globalization
Imports System.Collections.Generic
Imports NXOpen
Imports NXOpen.Assemblies

' NX 2306 target - VB.NET journal, native NX session.
' RUN: Tools > Journal > Play. First run with AUDIT_ONLY=True.
' Make the intended top assembly BOTH work and displayed part before running.
' Fully load exact geometry; choose the required arrangement/reference sets.
' Audit reports the loaded structure, not a geometric preview of simplification.
' Then set AUDIT_ONLY=False and run again to create a new output part.
' Default output folder: Documents\NX_Customer_Models. No existing file overwritten.
'
' REVIEW OF ORIGINAL:
' 1. .NET uses DestinationPart property, not C++ SetDestinationPart method.
' 2. Original swallowed load failures and could report incomplete output as success.
' 3. Original forced mm, ignored source-owned solids and had no rollback.
' 4. This revision audits every unsuppressed nested occurrence and selects each
'    top-level branch once. NX handles occurrence positions and nested structure;
'    do not copy prototype solids directly (that loses occurrence placement).
' 5. Output is flattened simplified SOLIDS, not a retained BOM or surface skin.
'    Enclosed body/void removal is not a guarantee of intellectual-property removal.
'
' API BASIS: Siemens-generated SimplifyBuilder reference (NX1953+), .NET
' declarations available for NX2212, and established PartCollection APIs.
' NX2306-specific compilation/runtime and geometry results HAVE NOT been tested here.
' Requires the applicable NX Assemblies license. No additional DLL is supplied.
' Teamcenter creation/export needs your site's managed FileNew workflow and is
' intentionally blocked for @DB source paths by this native-file implementation.
'
' TEST ON COPIES: nested/repeated rotated occurrences; source-owned solids;
' mm/inch source; suppressed branches; missing component; restrictive reference
' set; all geometry removed; existing output; save failure; undo after apply.
' Check output positions, envelope, interfaces, body count and section cuts.
' Inspect remaining internal details and feature links before customer release.
'
Module NX2306_3D_Assembly_Simplifier
    Private Const AUDIT_ONLY As Boolean = True
    Private Const OUTPUT_FOLDER As String = "" ' blank = Documents\NX_Customer_Models
    Private Const REMOVE_INTERNAL_BODIES As Boolean = True
    Private Const REMOVE_INTERNAL_VOIDS As Boolean = True
    Private Const UNITE_BODIES As Boolean = True
    ' Keep interface geometry by default. Enable only after engineering review.
    Private Const REMOVE_SMALL_BODIES As Boolean = False
    Private Const REMOVE_SMALL_HOLES As Boolean = False
    Private Const REMOVE_SMALL_BLENDS As Boolean = False
    Private Const MIN_BODY_MM As Double = 3.0
    Private Const MAX_HOLE_MM As Double = 5.0
    Private Const MAX_BLEND_MM As Double = 2.0

    Private sess As Session
    Private log As ListingWindow
    Private occurrenceCount As Integer
    Private suppressedCount As Integer
    Private errorCount As Integer
    Private solidCount As Integer
    Private sheetCount As Integer

    Public Sub Main()
        sess = Session.GetSession()
        log = sess.ListingWindow
        log.Open()
        occurrenceCount = 0 : suppressedCount = 0 : errorCount = 0
        solidCount = 0 : sheetCount = 0
        Dim mark As Session.UndoMarkId
        Dim hasMark As Boolean = False
        Dim saveStarted As Boolean = False
        Dim outputPath As String = ""
        Try
            Say("NX 2306 target | 3D assembly simplifier | " & DateTime.Now.ToString("s"))
            Say("Loaded NXOpen assembly: " & GetType(Session).Assembly.FullName)
            Say("AUDIT_ONLY=" & AUDIT_ONLY.ToString())
            Dim source As Part = sess.Parts.Work
            If source Is Nothing Then Throw New InvalidOperationException("Open a 3D assembly first.")
            If sess.Parts.Display Is Nothing OrElse sess.Parts.Display.Tag <> source.Tag Then
                Throw New InvalidOperationException("Make the intended top assembly both work and displayed part.")
            End If
            Say("Source: " & source.FullPath & " | Units: " & source.PartUnits.ToString())
            Dim root As Component = source.ComponentAssembly.RootComponent
            If root Is Nothing Then Throw New InvalidOperationException("The work part has no assembly root.")

            ' In audit mode no parts are loaded and no model changes are made.
            ' Apply mode requests full load and treats any reported load error as fatal.
            If Not AUDIT_ONLY Then FullyLoad(source)
            Dim selections As New List(Of DisplayableObject)()
            For Each body As Body In source.Bodies
                If body.IsSolidBody Then selections.Add(body)
            Next
            InspectBodies(source, "ROOT")
            For Each child As Component In root.GetChildren()
                Try
                    If Not child.IsSuppressed Then selections.Add(child)
                    AuditBranch(child, "ROOT", 1)
                Catch ex As Exception
                    errorCount += 1
                    Say("ERROR: top-level component: " & ex.Message)
                End Try
            Next
            Say("Occurrences=" & occurrenceCount.ToString() &
                "; suppressed branches=" & suppressedCount.ToString() &
                "; prototype solids counted per occurrence=" & solidCount.ToString() &
                "; sheets=" & sheetCount.ToString() & "; errors=" & errorCount.ToString())
            Say("Scope: current arrangement/reference sets; all selected branches, including hidden components.")
            Say("Counts describe prototypes, not guaranteed reference-set-visible output geometry.")
            Say("Remove internal bodies=" & REMOVE_INTERNAL_BODIES.ToString() &
                "; fill enclosed voids=" & REMOVE_INTERNAL_VOIDS.ToString() &
                "; unite overlapping bodies=" & UNITE_BODIES.ToString())
            Say("Small-body filter=" & REMOVE_SMALL_BODIES.ToString() & " (" & MIN_BODY_MM.ToString(CultureInfo.InvariantCulture) & " mm box diagonal)")
            Say("Hole filter=" & REMOVE_SMALL_HOLES.ToString() & " (" & MAX_HOLE_MM.ToString(CultureInfo.InvariantCulture) & " mm diameter)")
            Say("Blend filter=" & REMOVE_SMALL_BLENDS.ToString() & " (" & MAX_BLEND_MM.ToString(CultureInfo.InvariantCulture) & " mm radius)")
            If AUDIT_ONLY Then
                Say("AUDIT COMPLETE. Load/resolve reported components, review scope, then set AUDIT_ONLY=False.")
                Return
            End If
            If errorCount > 0 Then Throw New InvalidOperationException("Assembly audit failed. No output generated.")
            If selections.Count = 0 OrElse solidCount = 0 Then Throw New InvalidOperationException("No eligible solid geometry found.")
            If source.FullPath.StartsWith("@DB", StringComparison.OrdinalIgnoreCase) Then
                Throw New InvalidOperationException("Managed source detected. This version creates native files only; use a native exported assembly or a site-specific managed destination workflow.")
            End If
            outputPath = MakeOutputPath(source)
            Say("Output: " & outputPath)
            mark = sess.SetUndoMark(Session.MarkVisibility.Visible, "Create customer simplified assembly")
            hasMark = True
            Dim destination As Part = CreateDestination(source, outputPath)
            If sess.Parts.Work.Tag <> source.Tag OrElse sess.Parts.Display.Tag <> source.Tag Then
                Throw New InvalidOperationException("Part creation changed source context unexpectedly; stopping.")
            End If
            Simplify(source, destination, selections)
            ' Builder must be destroyed before inspecting the final output.
            Dim outputSolids As Integer = 0
            For Each body As Body In destination.Bodies
                If body.IsSolidBody Then outputSolids += 1
            Next
            If outputSolids = 0 Then Throw New InvalidOperationException("Simplification returned no solid bodies. Nothing will be saved.")
            Say("Result solid bodies=" & outputSolids.ToString())
            ' Reject a file created by another process since MakeOutputPath.
            If File.Exists(outputPath) Then Throw New IOException("Output path now exists; refusing to overwrite it.")
            saveStarted = True
            Dim saveStatus As PartSaveStatus = Nothing
            Try
                saveStatus = destination.Save(BasePart.SaveComponents.False, BasePart.CloseAfterSave.False)
                If saveStatus IsNot Nothing Then
                    If saveStatus.NumberUnsavedParts > 0 OrElse saveStatus.NumberUnsavedObjects > 0 Then
                        Throw New IOException("NX reported unsaved parts/objects. Review output; success is not confirmed.")
                    End If
                End If
            Finally
                If saveStatus IsNot Nothing Then saveStatus.Dispose()
            End Try
            If Not File.Exists(outputPath) Then Throw New IOException("NX returned from Save, but the output file was not found.")
            Say("CREATED: " & outputPath)
            Say("Source was not saved. Output remains loaded; open/display it for review.")
            Say("ENGINEERING REVIEW: section the result; check interfaces and all retained internal geometry.")
            Say("This is a simplified solid part, not a surface-only skin or a guaranteed sanitized customer export.")
        Catch ex As Exception
            Say("FAILED: " & ex.GetType().Name & ": " & ex.Message)
            If hasMark AndAlso Not saveStarted Then
                Try
                    sess.UndoToMark(mark, Nothing)
                    Say("Rolled back model operations to the undo mark.")
                Catch undoEx As Exception
                    Say("ROLLBACK FAILED: " & undoEx.Message & ". Inspect session state before continuing.")
                End Try
            End If
            If saveStarted Then Say("Save was attempted. Output/session retained for inspection: " & outputPath)
            Say("Loading and filesystem effects are not undone. No source Save or SaveAll was called.")
        End Try
    End Sub

    Private Sub FullyLoad(ByVal source As Part)
        Dim status As PartLoadStatus = Nothing
        Try
            status = sess.Parts.EnsurePartsLoadedFully(New BasePart() {source}, True)
            If status IsNot Nothing AndAlso status.NumberUnloadedParts > 0 Then
                For i As Integer = 0 To status.NumberUnloadedParts - 1
                    Say("LOAD ERROR: " & status.GetPartName(i) & " | " & status.GetStatusDescription(i))
                Next
                Throw New InvalidOperationException("Full assembly loading reported failures.")
            End If
        Finally
            If status IsNot Nothing Then status.Dispose()
        End Try
    End Sub

    Private Sub AuditBranch(ByVal component As Component, ByVal parentPath As String, ByVal depth As Integer)
        Dim label As String = parentPath & "/<unresolved>"
        Try
            If depth > 256 Then Throw New InvalidOperationException("Structure exceeds supported audit depth (256).")
            label = parentPath & "/" & component.DisplayName
            If component.IsSuppressed Then
                suppressedCount += 1
                Say("SKIP suppressed branch: " & label)
                Return
            End If
            occurrenceCount += 1
            Say("COMPONENT: " & label & " | refset=" & component.ReferenceSet)
            Dim prototype As Part = TryCast(component.Prototype, Part)
            If prototype Is Nothing Then Throw New InvalidOperationException("Exact prototype unavailable; fully load component.")
            If Not prototype.IsFullyLoaded Then Throw New InvalidOperationException("Prototype is not fully loaded.")
            InspectBodies(prototype, label)
            For Each child As Component In component.GetChildren()
                AuditBranch(child, label, depth + 1)
            Next
        Catch ex As Exception
            errorCount += 1
            Say("ERROR: " & label & " | " & ex.Message)
        End Try
    End Sub

    Private Sub InspectBodies(ByVal part As Part, ByVal label As String)
        For Each body As Body In part.Bodies
            Try
                If body.IsSolidBody Then
                    solidCount += 1
                Else
                    sheetCount += 1
                    Say("NOTICE: non-solid body in " & label & "; surface-only geometry is outside this tool's scope.")
                End If
            Catch ex As Exception
                errorCount += 1
                Say("ERROR: body inspection in " & label & " | " & ex.Message)
            End Try
        Next
    End Sub

    Private Function MakeOutputPath(ByVal source As Part) As String
        Dim folder As String = OUTPUT_FOLDER
        If String.IsNullOrWhiteSpace(folder) Then
            folder = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), "NX_Customer_Models")
        End If
        If Not Path.IsPathRooted(folder) Then Throw New IOException("OUTPUT_FOLDER must be an absolute path.")
        Directory.CreateDirectory(folder)
        Dim name As String = Path.GetFileNameWithoutExtension(source.Leaf)
        For Each c As Char In Path.GetInvalidFileNameChars()
            name = name.Replace(c, "_"c)
        Next
        If String.IsNullOrWhiteSpace(name) Then name = "Assembly"
        Dim target As String = Path.Combine(folder, name & "_SIMPLIFIED_" & DateTime.Now.ToString("yyyyMMdd_HHmmss") & "_" & Guid.NewGuid().ToString("N").Substring(0, 8) & ".prt")
        If File.Exists(target) Then Throw New IOException("Output already exists.")
        Return target
    End Function

    Private Function CreateDestination(ByVal source As Part, ByVal target As String) As Part
        Dim builder As FileNew = sess.Parts.FileNew()
        Dim created As NXObject = Nothing
        Try
            builder.UseBlankTemplate = True
            builder.ApplicationName = "Modeling"
            builder.Units = source.PartUnits
            builder.NewFileName = target
            builder.MakeDisplayedPart = False
            created = builder.Commit()
        Finally
            builder.Destroy()
        End Try
        Dim destination As Part = TryCast(created, Part)
        If destination Is Nothing Then destination = TryCast(sess.Parts.FindObject(target), Part)
        If destination Is Nothing Then Throw New InvalidOperationException("New destination part was not found.")
        If destination.Tag = source.Tag Then Throw New InvalidOperationException("Source and destination must differ.")
        Return destination
    End Function

    Private Sub Simplify(ByVal source As Part, ByVal destination As Part, ByVal selections As List(Of DisplayableObject))
        Dim builder As SimplifyBuilder = Nothing
        Try
            builder = source.AssemblyManager.CreateSimplifyBuilder()
            builder.DestinationPart = destination
            For Each item As DisplayableObject In selections
                ' Any selection failure is fatal: never silently omit an assembly branch.
                If Not builder.ObjectsToSimplify.Add(item) Then
                    Throw New InvalidOperationException("Could not select object tag " & item.Tag.ToString())
                End If
            Next
            builder.RemoveInternalBodies = REMOVE_INTERNAL_BODIES
            builder.RemoveInternalVoids = REMOVE_INTERNAL_VOIDS
            builder.UniteBodies = UNITE_BODIES
            builder.UseBodySize = REMOVE_SMALL_BODIES
            builder.UseHoleDiameter = REMOVE_SMALL_HOLES
            builder.UseBlendRadius = REMOVE_SMALL_BLENDS
            ' Explicit physical units avoid locale and source-unit ambiguity.
            If REMOVE_SMALL_BODIES Then builder.MinimumBodySize.RightHandSide = Mm(MIN_BODY_MM)
            If REMOVE_SMALL_HOLES Then builder.MaximumHoleSize.RightHandSide = Mm(MAX_HOLE_MM)
            If REMOVE_SMALL_BLENDS Then builder.MaximumBlendRadius.RightHandSide = Mm(MAX_BLEND_MM)
            If Not builder.Validate() Then Throw New InvalidOperationException("Simplify builder validation failed.")
            Say("Simplifying selected 3D structure...")
            builder.Commit()
        Finally
            If builder IsNot Nothing Then builder.Destroy()
        End Try
    End Sub

    Private Function Mm(ByVal value As Double) As String
        If value <= 0.0 Then Throw New ArgumentOutOfRangeException("value", "Threshold must be positive.")
        Return value.ToString("R", CultureInfo.InvariantCulture) & " mm"
    End Function

    Private Sub Say(ByVal message As String)
        log.WriteLine(message)
    End Sub

    Public Function GetUnloadOption(ByVal dummy As String) As Integer
        Return CInt(Session.LibraryUnloadOption.Immediately)
    End Function
End Module
