Option Strict On

Imports System
Imports System.IO
Imports System.Text
Imports System.Globalization
Imports System.Collections.Generic
Imports System.Windows.Forms
Imports Microsoft.VisualBasic

Imports NXOpen
Imports NXOpen.Assemblies
Imports NXOpen.UF

'==============================================================================
' NX 2306 - MASS PROPERTIES & INERTIA REPORT
'
' OUTPUT UNITS
' ------------
' Mass                  : kg
' Gravitational weight  : N   (mass * STANDARD_GRAVITY)
' CG coordinates        : mm, relative to DISPLAY PART ABSOLUTE CSYS
' Volume                : mm^3
' Density               : kg/m^3 and kg/mm^3
' Moments of inertia    : kg*mm^2
' Products of inertia   : kg*mm^2
'
' PRODUCTS OF INERTIA CONVENTION USED IN THIS REPORT
' --------------------------------------------------
' Pxy = + Integral(x*y dm)
' Pxz = + Integral(x*z dm)
' Pyz = + Integral(y*z dm)
'
' Therefore the conventional inertia tensor used for principal-axis
' calculations is:
'
'       | Ixx   -Pxy  -Pxz |
'   I = | -Pxy  Iyy  -Pyz |
'       | -Pxz -Pyz   Izz |
'
' NX UF_MODL_ask_mass_props_3d returns "products of inertia" but the public
' UF documentation does not state the algebraic sign convention explicitly.
' This journal compares origin and centroidal products against the
' parallel-axis relationship and normalizes the reported values to the
' positive-integral convention above.
'
' ASSEMBLY HANDLING
' -----------------
' - Suppressed components are skipped.
' - Hidden/blanked components are INCLUDED by default because display state
'   normally should not alter physical mass. Change INCLUDE_HIDDEN_COMPONENTS
'   if required.
' - The journal attempts to fully load the displayed assembly first.
' - Repeated component occurrences are processed separately.
' - Standard occurrence bodies are obtained with Component.FindOccurrence().
' - If an occurrence body cannot be resolved (for example special promoted or
'   deformed-body cases), it is skipped with a diagnostic.
'
' DENSITY HANDLING
' ----------------
' - If a Physical Material is assigned to the prototype body, NX's body
'   density is used.
' - If no Physical Material is assigned, the user is prompted once for a
'   fallback density in kg/m^3.
' - The source body/material is NOT edited. The UF mass/inertia values are
'   scaled to the fallback density mathematically.
'
' IMPORTANT
' ---------
' This journal changes the DISPLAY PART WCS temporarily to absolute
' (origin 0,0,0 and identity orientation) because AskMassProps3d reports CG
' and inertia in WCS coordinates. The original WCS is restored in Finally.
'
'==============================================================================

Module NX2306_Mass_Properties_Inertia_Report

    '==========================================================================
    ' NX SESSION
    '==========================================================================

    Private theSession As Session
    Private theUfSession As UFSession
    Private theUI As UI
    Private modelPart As Part
    Private lw As ListingWindow

    '==========================================================================
    ' USER SETTINGS
    '==========================================================================

    Private Const STANDARD_GRAVITY As Double = 9.80665       'm/s^2
    Private Const UF_ACCURACY As Double = 0.999

    'True = hidden/blanked assembly components still contribute to mass.
    Private Const INCLUDE_HIDDEN_COMPONENTS As Boolean = True

    'If True, a row is printed for every active component occurrence.
    'Subassembly rows include the mass of their complete active subtree.
    Private Const REPORT_EACH_COMPONENT_OCCURRENCE As Boolean = True

    'Part number aliases - customize for your Teamcenter / NX environment.
    Private ReadOnly PART_NUMBER_ALIASES As String() = {
        "DB_PART_NO",
        "DB_PART_NUMBER",
        "PART_NUMBER",
        "PART NO",
        "PART_NO",
        "ITEM_ID",
        "ITEM ID"
    }

    '==========================================================================
    ' GLOBAL STATE
    '==========================================================================

    Private reportRows As New List(Of MassResult)

    Private warningCount As Integer = 0
    Private errorCount As Integer = 0

    Private suppressedComponentCount As Integer = 0
    Private hiddenComponentCount As Integer = 0
    Private unresolvedBodyCount As Integer = 0
    Private skippedNonSolidCount As Integer = 0

    Private fallbackDensityKgM3 As Double = Double.NaN
    Private fallbackDensityPrompted As Boolean = False

    Private csvFolder As String = ""
    Private csvPath As String = ""

    'Saved WCS state.
    Private oldWcsOrigin As Point3d
    Private oldWcsMatrix As Matrix3x3
    Private wcsWasChanged As Boolean = False

    '==========================================================================
    ' TYPES
    '==========================================================================

    Private Enum RunMode
        CompleteModel
        SelectedBodies
    End Enum

    Private Class BodyItem
        Public OccurrenceBody As Body
        Public PrototypeBody As Body
        Public Component As Component
        Public PrototypePart As Part

        Public PartName As String = ""
        Public ComponentName As String = ""
        Public ComponentPath As String = ""
        Public PartNumber As String = ""
        Public Hidden As Boolean = False
    End Class

    Private Class MassResult
        Public RowType As String = ""

        Public PartName As String = ""
        Public ComponentName As String = ""
        Public ComponentPath As String = ""
        Public PartNumber As String = ""

        Public MaterialName As String = ""
        Public DensitySource As String = ""
        Public UsedAssumedDensity As Boolean = False
        Public Hidden As Boolean = False

        Public MassKg As Double = 0.0
        Public WeightN As Double = 0.0
        Public VolumeMm3 As Double = 0.0
        Public DensityKgM3 As Double = 0.0
        Public DensityKgMm3 As Double = 0.0

        Public CgX As Double = 0.0
        Public CgY As Double = 0.0
        Public CgZ As Double = 0.0

        'About absolute origin.
        Public IxxOrigin As Double = 0.0
        Public IyyOrigin As Double = 0.0
        Public IzzOrigin As Double = 0.0
        Public PxyOrigin As Double = 0.0
        Public PxzOrigin As Double = 0.0
        Public PyzOrigin As Double = 0.0

        'About combined CG, axes parallel to absolute axes.
        Public IxxCg As Double = 0.0
        Public IyyCg As Double = 0.0
        Public IzzCg As Double = 0.0
        Public PxyCg As Double = 0.0
        Public PxzCg As Double = 0.0
        Public PyzCg As Double = 0.0

        'Principal moments and directions relative to absolute XYZ.
        Public Principal1 As Double = 0.0
        Public Principal2 As Double = 0.0
        Public Principal3 As Double = 0.0

        Public Axis1X As Double = 1.0
        Public Axis1Y As Double = 0.0
        Public Axis1Z As Double = 0.0

        Public Axis2X As Double = 0.0
        Public Axis2Y As Double = 1.0
        Public Axis2Z As Double = 0.0

        Public Axis3X As Double = 0.0
        Public Axis3Y As Double = 0.0
        Public Axis3Z As Double = 1.0

        Public Notes As String = ""
    End Class

    '==========================================================================
    ' MAIN
    '==========================================================================

    Public Sub Main()

        theSession = Session.GetSession()
        theUfSession = UFSession.GetUFSession()
        theUI = UI.GetUI()
        lw = theSession.ListingWindow
        lw.Open()

        modelPart = TryCast(theSession.Parts.Display, Part)

        PrintHeader()

        If modelPart Is Nothing Then
            LogError("No displayed NX Part is available.")
            Return
        End If

        Dim mode As RunMode

        If Not AskRunMode(mode) Then
            lw.WriteLine("Cancelled by user.")
            Return
        End If

        If Not AskCsvFolder() Then
            lw.WriteLine("CSV folder selection cancelled. Journal stopped.")
            Return
        End If

        Try

            'Fully load displayed assembly / part hierarchy.
            EnsureDisplayedModelFullyLoaded()

            'AskMassProps3d reports values in WCS.
            'Temporarily make WCS = absolute CSYS.
            SetWcsToAbsolute()

            reportRows.Clear()

            If mode = RunMode.SelectedBodies Then

                ProcessSelectedBodies()

            Else

                ProcessCompleteModel()

            End If

            If reportRows.Count = 0 Then
                LogError("No valid solid-body mass properties were produced.")
                Return
            End If

            PrintAllResults()

            ExportCsv()

            PrintRunSummary()

        Catch ex As Exception

            LogError(ex.Message)

            If Not String.IsNullOrWhiteSpace(ex.StackTrace) Then
                lw.WriteLine(ex.StackTrace)
            End If

        Finally

            RestoreOriginalWcs()

        End Try

    End Sub

    '==========================================================================
    ' USER CHOICES
    '==========================================================================

    Private Function AskRunMode(
        ByRef mode As RunMode) As Boolean

        Dim answer As DialogResult =
            MessageBox.Show(
                "Mass Properties Report" & Environment.NewLine &
                Environment.NewLine &
                "YES  = Process complete displayed part / assembly" &
                Environment.NewLine &
                "NO   = Select solid bodies" &
                Environment.NewLine &
                "CANCEL = Exit",
                "NX 2306 Mass Properties",
                MessageBoxButtons.YesNoCancel,
                MessageBoxIcon.Question)

        Select Case answer

            Case DialogResult.Yes
                mode = RunMode.CompleteModel
                Return True

            Case DialogResult.No
                mode = RunMode.SelectedBodies
                Return True

            Case Else
                Return False

        End Select

    End Function

    Private Function AskCsvFolder() As Boolean

        Using dlg As New FolderBrowserDialog()

            dlg.Description =
                "Select folder for NX mass properties CSV report"

            dlg.ShowNewFolderButton = True

            Try
                If Not String.IsNullOrWhiteSpace(modelPart.FullPath) Then
                    Dim initialFolder As String =
                        Path.GetDirectoryName(modelPart.FullPath)

                    If Directory.Exists(initialFolder) Then
                        dlg.SelectedPath = initialFolder
                    End If
                End If
            Catch
            End Try

            If dlg.ShowDialog() <> DialogResult.OK Then
                Return False
            End If

            csvFolder = dlg.SelectedPath

        End Using

        Return True

    End Function

    '==========================================================================
    ' FULL LOAD
    '==========================================================================

    Private Sub EnsureDisplayedModelFullyLoaded()

        Dim parts(0) As BasePart
        parts(0) = modelPart

        Dim loadStatus As PartLoadStatus = Nothing

        Try

            loadStatus =
                theSession.Parts.EnsurePartsLoadedFully(
                    parts,
                    True)

            LogInfo("Displayed model load: full-load request completed.")

        Finally

            If loadStatus IsNot Nothing Then
                loadStatus.Dispose()
            End If

        End Try

    End Sub

    Private Function EnsurePrototypePartLoaded(
        ByVal part As Part) As Boolean

        If part Is Nothing Then Return False

        If part.IsFullyLoaded Then Return True

        Dim status As PartLoadStatus = Nothing

        Try

            status = part.LoadThisPartFully()

            Return part.IsFullyLoaded

        Catch ex As Exception

            LogWarning(
                "Unable to fully load part '" &
                part.Leaf &
                "': " &
                ex.Message)

            Return False

        Finally

            If status IsNot Nothing Then
                status.Dispose()
            End If

        End Try

    End Function

    '==========================================================================
    ' WCS -> ABSOLUTE
    '==========================================================================

    Private Sub SetWcsToAbsolute()

        oldWcsOrigin =
            modelPart.WCS.Origin

        oldWcsMatrix =
            modelPart.WCS.CoordinateSystem.Orientation.Element

        Dim absoluteOrigin As New Point3d(0.0, 0.0, 0.0)

        Dim identity As Matrix3x3 =
            IdentityMatrix()

        modelPart.WCS.SetOriginAndMatrix(
            absoluteOrigin,
            identity)

        wcsWasChanged = True

        LogInfo(
            "Mass-property coordinate system: DISPLAY PART ABSOLUTE CSYS.")

    End Sub

    Private Sub RestoreOriginalWcs()

        If Not wcsWasChanged Then Return

        Try

            modelPart.WCS.SetOriginAndMatrix(
                oldWcsOrigin,
                oldWcsMatrix)

            wcsWasChanged = False

            LogInfo("Original WCS restored.")

        Catch ex As Exception

            LogWarning(
                "Unable to restore the original WCS: " &
                ex.Message)

        End Try

    End Sub

    Private Function IdentityMatrix() As Matrix3x3

        Dim m As New Matrix3x3()

        m.Xx = 1.0 : m.Xy = 0.0 : m.Xz = 0.0
        m.Yx = 0.0 : m.Yy = 1.0 : m.Yz = 0.0
        m.Zx = 0.0 : m.Zy = 0.0 : m.Zz = 1.0

        Return m

    End Function

    '==========================================================================
    ' COMPLETE MODEL
    '==========================================================================

    Private Sub ProcessCompleteModel()

        Dim root As Component =
            modelPart.ComponentAssembly.RootComponent

        Dim topChildren() As Component = Nothing

        If root IsNot Nothing Then
            Try
                topChildren = root.GetChildren()
            Catch
                topChildren = Nothing
            End Try
        End If

        Dim isAssembly As Boolean =
            (topChildren IsNot Nothing AndAlso topChildren.Length > 0)

        If Not isAssembly Then

            ProcessSinglePart()

            Return

        End If

        LogInfo("")
        LogInfo("Processing complete assembly...")

        Dim totalMembers As New List(Of MassResult)

        'Bodies owned directly by the assembly part itself.
        For Each body As Body In modelPart.Bodies

            If Not SafeIsSolid(body) Then
                skippedNonSolidCount += 1
                Continue For
            End If

            Dim item As BodyItem =
                CreateRootBodyItem(body)

            Dim bodyResult As MassResult =
                MeasureBodyItem(item)

            If bodyResult IsNot Nothing Then
                totalMembers.Add(bodyResult)
            End If

        Next

        'Each top-level component is recursively aggregated.
        For Each child As Component In topChildren

            Dim childResult As MassResult =
                ProcessComponentRecursive(child)

            If childResult IsNot Nothing AndAlso
               childResult.MassKg > 0.0 Then

                totalMembers.Add(childResult)

            End If

        Next

        Dim assemblyTotal As MassResult =
            CombineResults(
                totalMembers,
                "ASSEMBLY TOTAL",
                modelPart.Leaf,
                "<ROOT ASSEMBLY>",
                modelPart.Leaf,
                FindPartNumber(Nothing, modelPart),
                False)

        If assemblyTotal IsNot Nothing Then
            assemblyTotal.MaterialName = "MIXED / ASSEMBLY"
            reportRows.Add(assemblyTotal)
        End If

    End Sub

    '==========================================================================
    ' SINGLE PART
    '==========================================================================

    Private Sub ProcessSinglePart()

        LogInfo("")
        LogInfo("Processing complete part...")

        Dim bodyResults As New List(Of MassResult)

        For Each body As Body In modelPart.Bodies

            If Not SafeIsSolid(body) Then
                skippedNonSolidCount += 1
                Continue For
            End If

            Dim item As BodyItem =
                CreateRootBodyItem(body)

            Dim r As MassResult =
                MeasureBodyItem(item)

            If r IsNot Nothing Then
                bodyResults.Add(r)
            End If

        Next

        Dim total As MassResult =
            CombineResults(
                bodyResults,
                "PART TOTAL",
                modelPart.Leaf,
                "<WORK PART>",
                modelPart.Leaf,
                FindPartNumber(Nothing, modelPart),
                False)

        If total IsNot Nothing Then

            total.MaterialName =
                AggregateMaterialLabel(bodyResults)

            reportRows.Add(total)

        End If

    End Sub

    '==========================================================================
    ' RECURSIVE ASSEMBLY COMPONENT
    '
    ' Returned result includes:
    '   - bodies owned directly by the component prototype part
    '   - all active child-component subtrees
    '
    ' This makes a subassembly occurrence row represent the complete
    ' subassembly occurrence mass.
    '==========================================================================

    Private Function ProcessComponentRecursive(
        ByVal component As Component) As MassResult

        If component Is Nothing Then
            Return Nothing
        End If

        Try
            If component.IsSuppressed Then

                suppressedComponentCount += 1

                LogInfo(
                    "SKIP SUPPRESSED: " &
                    ComponentPath(component))

                Return Nothing

            End If
        Catch ex As Exception

            LogWarning(
                "Could not query suppression state for " &
                SafeComponentName(component) &
                ": " &
                ex.Message)

        End Try

        Dim hidden As Boolean = False

        Try
            hidden = component.IsBlanked
        Catch
        End Try

        If hidden Then

            hiddenComponentCount += 1

            If Not INCLUDE_HIDDEN_COMPONENTS Then

                LogInfo(
                    "SKIP HIDDEN: " &
                    ComponentPath(component))

                Return Nothing

            End If

        End If

        Dim protoPart As Part =
            TryCast(component.Prototype, Part)

        If protoPart Is Nothing Then

            LogWarning(
                "Component prototype is not a loaded Part: " &
                ComponentPath(component))

            Return Nothing

        End If

        If Not EnsurePrototypePartLoaded(protoPart) Then

            LogWarning(
                "Skipping unloaded component: " &
                ComponentPath(component))

            Return Nothing

        End If

        Dim memberResults As New List(Of MassResult)

        '----------------------------------------------------------------------
        ' Direct bodies in this component occurrence.
        '----------------------------------------------------------------------

        For Each protoBody As Body In protoPart.Bodies

            If Not SafeIsSolid(protoBody) Then
                skippedNonSolidCount += 1
                Continue For
            End If

            Dim occurrenceObject As NXObject = Nothing

            Try
                occurrenceObject =
                    component.FindOccurrence(protoBody)
            Catch ex As Exception
                LogWarning(
                    "FindOccurrence failed: " &
                    ComponentPath(component) &
                    " | " &
                    ex.Message)
            End Try

            Dim occurrenceBody As Body =
                TryCast(occurrenceObject, Body)

            If occurrenceBody Is Nothing Then

                unresolvedBodyCount += 1

                LogWarning(
                    "No standard occurrence body resolved for prototype Body Tag " &
                    protoBody.Tag.ToString() &
                    " in component " &
                    ComponentPath(component) &
                    ". Promoted/deformed/reference-set special cases may require " &
                    "site-specific handling.")

                Continue For

            End If

            Dim item As New BodyItem()

            item.OccurrenceBody = occurrenceBody
            item.PrototypeBody = protoBody
            item.Component = component
            item.PrototypePart = protoPart
            item.PartName = protoPart.Leaf
            item.ComponentName = SafeComponentName(component)
            item.ComponentPath = ComponentPath(component)
            item.PartNumber = FindPartNumber(component, protoPart)
            item.Hidden = hidden

            Dim bodyResult As MassResult =
                MeasureBodyItem(item)

            If bodyResult IsNot Nothing Then
                memberResults.Add(bodyResult)
            End If

        Next

        '----------------------------------------------------------------------
        ' Child subassemblies / components.
        '----------------------------------------------------------------------

        Try

            For Each child As Component In component.GetChildren()

                Dim childResult As MassResult =
                    ProcessComponentRecursive(child)

                If childResult IsNot Nothing AndAlso
                   childResult.MassKg > 0.0 Then

                    memberResults.Add(childResult)

                End If

            Next

        Catch ex As Exception

            LogWarning(
                "Unable to traverse children of " &
                ComponentPath(component) &
                ": " &
                ex.Message)

        End Try

        If memberResults.Count = 0 Then
            Return Nothing
        End If

        Dim result As MassResult =
            CombineResults(
                memberResults,
                "COMPONENT",
                protoPart.Leaf,
                SafeComponentName(component),
                ComponentPath(component),
                FindPartNumber(component, protoPart),
                hidden)

        If result Is Nothing Then
            Return Nothing
        End If

        result.MaterialName =
            AggregateMaterialLabel(memberResults)

        If REPORT_EACH_COMPONENT_OCCURRENCE Then
            reportRows.Add(result)
        End If

        Return result

    End Function

    '==========================================================================
    ' SELECTED BODIES
    '==========================================================================

    Private Sub ProcessSelectedBodies()

        LogInfo("")
        LogInfo("Select one or more solid bodies...")

        Dim selected() As TaggedObject = Nothing

        Dim bodyMask As New Selection.MaskTriple(
            UFConstants.UF_solid_type,
            0,
            UFConstants.UF_UI_SEL_FEATURE_BODY)

        Dim masks() As Selection.MaskTriple = {
            bodyMask
        }

        Dim response As Selection.Response =
            theUI.SelectionManager.SelectTaggedObjects(
                "Select solid bodies for mass properties",
                "Mass Properties - Select Bodies",
                Selection.SelectionScope.AnyInAssembly,
                Selection.SelectionAction.ClearAndEnableSpecific,
                False,
                False,
                masks,
                selected)

        If response <> Selection.Response.Ok OrElse
           selected Is Nothing OrElse
           selected.Length = 0 Then

            Throw New OperationCanceledException(
                "Body selection cancelled or empty.")

        End If

        Dim totals As New List(Of MassResult)

        For Each tagged As TaggedObject In selected

            Dim body As Body =
                TryCast(tagged, Body)

            If body Is Nothing OrElse
               Not SafeIsSolid(body) Then

                skippedNonSolidCount += 1
                Continue For

            End If

            Dim item As BodyItem =
                CreateBodyItemFromSelectedBody(body)

            Dim bodyResult As MassResult =
                MeasureBodyItem(item)

            If bodyResult Is Nothing Then
                Continue For
            End If

            bodyResult.RowType = "SELECTED BODY"

            reportRows.Add(bodyResult)
            totals.Add(bodyResult)

        Next

        Dim selectedTotal As MassResult =
            CombineResults(
                totals,
                "SELECTED TOTAL",
                modelPart.Leaf,
                "<SELECTED BODIES>",
                modelPart.Leaf,
                "",
                False)

        If selectedTotal IsNot Nothing Then

            selectedTotal.MaterialName =
                AggregateMaterialLabel(totals)

            reportRows.Add(selectedTotal)

        End If

    End Sub

    '==========================================================================
    ' BODY ITEM HELPERS
    '==========================================================================

    Private Function CreateRootBodyItem(
        ByVal body As Body) As BodyItem

        Dim item As New BodyItem()

        item.OccurrenceBody = body
        item.PrototypeBody = body
        item.Component = Nothing
        item.PrototypePart = modelPart
        item.PartName = modelPart.Leaf
        item.ComponentName = "<WORK PART>"
        item.ComponentPath = modelPart.Leaf
        item.PartNumber = FindPartNumber(Nothing, modelPart)
        item.Hidden = SafeIsBlanked(body)

        Return item

    End Function

    Private Function CreateBodyItemFromSelectedBody(
        ByVal body As Body) As BodyItem

        Dim item As New BodyItem()

        item.OccurrenceBody = body

        Dim protoBody As Body = body
        Dim comp As Component = Nothing

        Try
            comp = body.OwningComponent
        Catch
        End Try

        Try
            If body.IsOccurrence Then

                Dim pb As Body =
                    TryCast(body.Prototype, Body)

                If pb IsNot Nothing Then
                    protoBody = pb
                End If

            End If
        Catch
        End Try

        Dim protoPart As Part =
            TryCast(protoBody.OwningPart, Part)

        If protoPart Is Nothing Then
            protoPart = modelPart
        End If

        item.PrototypeBody = protoBody
        item.Component = comp
        item.PrototypePart = protoPart
        item.PartName = protoPart.Leaf
        item.ComponentName =
            If(comp Is Nothing, "<WORK PART>", SafeComponentName(comp))

        item.ComponentPath =
            If(comp Is Nothing, modelPart.Leaf, ComponentPath(comp))

        item.PartNumber =
            FindPartNumber(comp, protoPart)

        item.Hidden =
            If(comp Is Nothing, SafeIsBlanked(body), SafeIsBlanked(comp))

        Return item

    End Function

    '==========================================================================
    ' MASS PROPERTIES FOR ONE BODY / OCCURRENCE
    '==========================================================================

    Private Function MeasureBodyItem(
        ByVal item As BodyItem) As MassResult

        If item Is Nothing OrElse
           item.OccurrenceBody Is Nothing Then

            Return Nothing
        End If

        Dim tags() As Tag = {
            item.OccurrenceBody.Tag
        }

        Dim accuracyValues(10) As Double
        accuracyValues(0) = UF_ACCURACY

        Dim mp(46) As Double
        Dim stats(12) As Double

        Try

            'type = 1 : Solid bodies
            'units = 4: Kilograms and meters
            'density input is not used for solid-body analysis.
            'accuracy = 1: use accuracyValues(0)
            theUfSession.Modl.AskMassProps3d(
                tags,
                1,
                1,
                4,
                0.0,
                1,
                accuracyValues,
                mp,
                stats)

        Catch ex As Exception

            LogError(
                "Mass property calculation failed for " &
                item.ComponentPath &
                " | Body Tag " &
                item.OccurrenceBody.Tag.ToString() &
                ": " &
                ex.Message)

            Return Nothing

        End Try

        Dim result As New MassResult()

        result.RowType = "BODY"
        result.PartName = item.PartName
        result.ComponentName = item.ComponentName
        result.ComponentPath = item.ComponentPath
        result.PartNumber = item.PartNumber
        result.Hidden = item.Hidden

        '----------------------------------------------------------------------
        ' UF output in units=4:
        '   Volume  = m^3
        '   Mass    = kg
        '   CG      = m
        '   Inertia = kg*m^2
        '   Density = kg/m^3
        '----------------------------------------------------------------------

        Dim rawMassKg As Double = mp(2)
        Dim rawDensityKgM3 As Double = mp(46)

        Dim materialName As String =
            GetAssignedMaterialName(
                item.PrototypePart,
                item.PrototypeBody)

        Dim bodyDensityKgM3 As Double =
            GetBodyDensityKgM3(
                item.PrototypeBody)

        Dim densityScale As Double = 1.0

        If Not String.IsNullOrWhiteSpace(materialName) Then

            result.MaterialName = materialName
            result.DensitySource =
                "NX ASSIGNED PHYSICAL MATERIAL / BODY DENSITY"

            result.UsedAssumedDensity = False

        ElseIf bodyDensityKgM3 > 0.0 Then

            'A body can carry a valid NX density even when no named
            'Physical Material object is assigned.
            result.MaterialName = "<NO NAMED PHYSICAL MATERIAL>"
            result.DensitySource =
                "NX BODY DENSITY"

            result.UsedAssumedDensity = False

            result.Notes =
                "No named Physical Material found; existing NX body density used."

        Else

            Dim assumedDensity As Double =
                GetFallbackDensityKgM3()

            'For a solid body UF mass properties use the body's stored density.
            'To preserve the source model we do not call SetDensity().
            'Instead, when UF returned a usable nonzero baseline density,
            'mass and inertia are scaled mathematically to the user's fallback.
            If rawDensityKgM3 <= 0.0 Then

                LogError(
                    "Body has no usable NX density and UF returned zero/negative " &
                    "baseline density. The journal will not modify Body.Density. " &
                    "Body skipped: " &
                    item.ComponentPath)

                Return Nothing

            End If

            densityScale =
                assumedDensity /
                rawDensityKgM3

            result.MaterialName = "<NO PHYSICAL MATERIAL>"
            result.DensitySource =
                "ASSUMED USER DENSITY"

            result.UsedAssumedDensity = True

            result.Notes =
                "Fallback density applied mathematically; source body unchanged."

        End If

        '----------------------------------------------------------------------
        ' Detect / normalize UF products of inertia.
        ' mp(16..18) = origin products
        ' mp(19..21) = centroidal products
        '
        ' Normalize to:
        '   Pxy = +Integral(x*y dm)
        '----------------------------------------------------------------------

        Dim rawPositiveConvention As Boolean =
            InferUfPositiveProductConvention(
                rawMassKg,
                mp(3),
                mp(4),
                mp(5),
                mp(16),
                mp(17),
                mp(18),
                mp(19),
                mp(20),
                mp(21))

        Dim productFactor As Double =
            If(rawPositiveConvention, 1.0, -1.0)

        If Not rawPositiveConvention Then

            result.Notes =
                AppendNote(
                    result.Notes,
                    "UF raw products were sign-normalized to +Integral(x*y dm).")

        End If

        '----------------------------------------------------------------------
        ' Scale density-dependent properties if user fallback density is used.
        ' Volume and CG do not scale with density.
        '----------------------------------------------------------------------

        result.VolumeMm3 =
            mp(1) * 1.0E+9

        result.MassKg =
            rawMassKg * densityScale

        result.WeightN =
            result.MassKg * STANDARD_GRAVITY

        result.CgX =
            mp(3) * 1000.0

        result.CgY =
            mp(4) * 1000.0

        result.CgZ =
            mp(5) * 1000.0

        'kg*m^2 -> kg*mm^2 = x 1E6
        Dim inertiaScale As Double =
            densityScale * 1.0E+6

        result.IxxOrigin = mp(9) * inertiaScale
        result.IyyOrigin = mp(10) * inertiaScale
        result.IzzOrigin = mp(11) * inertiaScale

        result.IxxCg = mp(12) * inertiaScale
        result.IyyCg = mp(13) * inertiaScale
        result.IzzCg = mp(14) * inertiaScale

        result.PxyOrigin =
            mp(16) * productFactor * inertiaScale

        result.PxzOrigin =
            mp(17) * productFactor * inertiaScale

        result.PyzOrigin =
            mp(18) * productFactor * inertiaScale

        result.PxyCg =
            mp(19) * productFactor * inertiaScale

        result.PxzCg =
            mp(20) * productFactor * inertiaScale

        result.PyzCg =
            mp(21) * productFactor * inertiaScale

        If result.VolumeMm3 > 0.0 Then

            result.DensityKgMm3 =
                result.MassKg / result.VolumeMm3

            result.DensityKgM3 =
                result.DensityKgMm3 * 1.0E+9

        Else

            result.DensityKgMm3 = 0.0
            result.DensityKgM3 = 0.0

        End If

        CalculatePrincipalProperties(result)

        Return result

    End Function

    '==========================================================================
    ' UF PRODUCT SIGN INFERENCE
    '
    ' Positive product convention:
    ' Pxy_O = Pxy_C + M*Cx*Cy
    '
    ' Negative/tensor-offdiagonal convention:
    ' Qxy_O = Qxy_C - M*Cx*Cy
    '
    ' Returns True for positive convention.
    ' If the CG offset is too small to distinguish conventions, the journal
    ' defaults to positive convention.
    '==========================================================================

    Private Function InferUfPositiveProductConvention(
        ByVal massKg As Double,
        ByVal xM As Double,
        ByVal yM As Double,
        ByVal zM As Double,
        ByVal pxyOrigin As Double,
        ByVal pxzOrigin As Double,
        ByVal pyzOrigin As Double,
        ByVal pxyCg As Double,
        ByVal pxzCg As Double,
        ByVal pyzCg As Double) As Boolean

        Dim shiftXY As Double = massKg * xM * yM
        Dim shiftXZ As Double = massKg * xM * zM
        Dim shiftYZ As Double = massKg * yM * zM

        Dim shiftMagnitude As Double =
            Math.Abs(shiftXY) +
            Math.Abs(shiftXZ) +
            Math.Abs(shiftYZ)

        If shiftMagnitude < 1.0E-12 Then
            Return True
        End If

        Dim positiveError As Double =
            Math.Abs(pxyOrigin - (pxyCg + shiftXY)) +
            Math.Abs(pxzOrigin - (pxzCg + shiftXZ)) +
            Math.Abs(pyzOrigin - (pyzCg + shiftYZ))

        Dim negativeError As Double =
            Math.Abs(pxyOrigin - (pxyCg - shiftXY)) +
            Math.Abs(pxzOrigin - (pxzCg - shiftXZ)) +
            Math.Abs(pyzOrigin - (pyzCg - shiftYZ))

        Return positiveError <= negativeError

    End Function

    '==========================================================================
    ' COMBINE MASS RESULTS
    '
    ' All member results are already expressed in the same absolute axes.
    ' Origin inertias/products can therefore be summed directly.
    '
    ' Combined CG:
    ' C = Sum(mi*Ci) / Sum(mi)
    '
    ' Parallel-axis relation for positive products:
    ' Pxy_CG = Pxy_O - M*Cx*Cy
    '==========================================================================

    Private Function CombineResults(
        ByVal members As List(Of MassResult),
        ByVal rowType As String,
        ByVal partName As String,
        ByVal componentName As String,
        ByVal componentPath As String,
        ByVal partNumber As String,
        ByVal hidden As Boolean) As MassResult

        If members Is Nothing OrElse members.Count = 0 Then
            Return Nothing
        End If

        Dim result As New MassResult()

        result.RowType = rowType
        result.PartName = partName
        result.ComponentName = componentName
        result.ComponentPath = componentPath
        result.PartNumber = partNumber
        result.Hidden = hidden

        Dim mx As Double = 0.0
        Dim my As Double = 0.0
        Dim mz As Double = 0.0

        For Each member As MassResult In members

            If member Is Nothing OrElse member.MassKg <= 0.0 Then
                Continue For
            End If

            result.MassKg += member.MassKg
            result.VolumeMm3 += member.VolumeMm3

            mx += member.MassKg * member.CgX
            my += member.MassKg * member.CgY
            mz += member.MassKg * member.CgZ

            result.IxxOrigin += member.IxxOrigin
            result.IyyOrigin += member.IyyOrigin
            result.IzzOrigin += member.IzzOrigin

            result.PxyOrigin += member.PxyOrigin
            result.PxzOrigin += member.PxzOrigin
            result.PyzOrigin += member.PyzOrigin

            If member.UsedAssumedDensity Then
                result.UsedAssumedDensity = True
            End If

        Next

        If result.MassKg <= 0.0 Then
            Return Nothing
        End If

        result.WeightN =
            result.MassKg * STANDARD_GRAVITY

        result.CgX = mx / result.MassKg
        result.CgY = my / result.MassKg
        result.CgZ = mz / result.MassKg

        'Origin -> combined CG using parallel-axis theorem.
        result.IxxCg =
            result.IxxOrigin -
            result.MassKg *
            (result.CgY * result.CgY +
             result.CgZ * result.CgZ)

        result.IyyCg =
            result.IyyOrigin -
            result.MassKg *
            (result.CgX * result.CgX +
             result.CgZ * result.CgZ)

        result.IzzCg =
            result.IzzOrigin -
            result.MassKg *
            (result.CgX * result.CgX +
             result.CgY * result.CgY)

        result.PxyCg =
            result.PxyOrigin -
            result.MassKg *
            result.CgX *
            result.CgY

        result.PxzCg =
            result.PxzOrigin -
            result.MassKg *
            result.CgX *
            result.CgZ

        result.PyzCg =
            result.PyzOrigin -
            result.MassKg *
            result.CgY *
            result.CgZ

        If result.VolumeMm3 > 0.0 Then

            result.DensityKgMm3 =
                result.MassKg /
                result.VolumeMm3

            result.DensityKgM3 =
                result.DensityKgMm3 *
                1.0E+9

        End If

        result.DensitySource =
            If(
                result.UsedAssumedDensity,
                "EFFECTIVE DENSITY; ASSUMPTION PRESENT",
                "EFFECTIVE DENSITY FROM NX MATERIAL/BODY DENSITIES")

        result.Notes =
            If(
                result.UsedAssumedDensity,
                "One or more members used user-assumed fallback density.",
                "")

        CalculatePrincipalProperties(result)

        Return result

    End Function

    '==========================================================================
    ' PRINCIPAL MOMENTS / AXES
    '
    ' Build symmetric inertia tensor about CG:
    '
    '       | Ixx   -Pxy  -Pxz |
    '   I = | -Pxy  Iyy  -Pyz |
    '       | -Pxz -Pyz   Izz |
    '
    ' Then solve eigenproblem by Jacobi rotations.
    '==========================================================================

    Private Sub CalculatePrincipalProperties(
        ByVal result As MassResult)

        Dim a(2, 2) As Double

        a(0, 0) = result.IxxCg
        a(1, 1) = result.IyyCg
        a(2, 2) = result.IzzCg

        a(0, 1) = -result.PxyCg
        a(1, 0) = a(0, 1)

        a(0, 2) = -result.PxzCg
        a(2, 0) = a(0, 2)

        a(1, 2) = -result.PyzCg
        a(2, 1) = a(1, 2)

        Dim eigenValues() As Double = Nothing
        Dim eigenVectors(,) As Double = Nothing

        JacobiEigen3x3(
            a,
            eigenValues,
            eigenVectors)

        result.Principal1 = eigenValues(0)
        result.Principal2 = eigenValues(1)
        result.Principal3 = eigenValues(2)

        'Eigenvectors are stored by column.
        result.Axis1X = eigenVectors(0, 0)
        result.Axis1Y = eigenVectors(1, 0)
        result.Axis1Z = eigenVectors(2, 0)

        result.Axis2X = eigenVectors(0, 1)
        result.Axis2Y = eigenVectors(1, 1)
        result.Axis2Z = eigenVectors(2, 1)

        result.Axis3X = eigenVectors(0, 2)
        result.Axis3Y = eigenVectors(1, 2)
        result.Axis3Z = eigenVectors(2, 2)

    End Sub

    Private Sub JacobiEigen3x3(
        ByVal input(,) As Double,
        ByRef eigenValues() As Double,
        ByRef eigenVectors(,) As Double)

        Dim a(2, 2) As Double
        Dim v(2, 2) As Double

        For i As Integer = 0 To 2
            For j As Integer = 0 To 2
                a(i, j) = input(i, j)
                v(i, j) = If(i = j, 1.0, 0.0)
            Next
        Next

        For iteration As Integer = 0 To 59

            Dim p As Integer = 0
            Dim q As Integer = 1
            Dim maxOff As Double =
                Math.Abs(a(0, 1))

            If Math.Abs(a(0, 2)) > maxOff Then
                p = 0 : q = 2
                maxOff = Math.Abs(a(0, 2))
            End If

            If Math.Abs(a(1, 2)) > maxOff Then
                p = 1 : q = 2
                maxOff = Math.Abs(a(1, 2))
            End If

            Dim diagScale As Double =
                Math.Max(
                    1.0,
                    Math.Max(
                        Math.Abs(a(0, 0)),
                        Math.Max(
                            Math.Abs(a(1, 1)),
                            Math.Abs(a(2, 2)))))

            If maxOff <= diagScale * 1.0E-12 Then
                Exit For
            End If

            Dim app As Double = a(p, p)
            Dim aqq As Double = a(q, q)
            Dim apq As Double = a(p, q)

            Dim phi As Double =
                0.5 *
                Math.Atan2(
                    2.0 * apq,
                    aqq - app)

            Dim c As Double = Math.Cos(phi)
            Dim s As Double = Math.Sin(phi)

            For k As Integer = 0 To 2

                If k = p OrElse k = q Then
                    Continue For
                End If

                Dim akp As Double = a(k, p)
                Dim akq As Double = a(k, q)

                Dim newKp As Double =
                    c * akp -
                    s * akq

                Dim newKq As Double =
                    s * akp +
                    c * akq

                a(k, p) = newKp
                a(p, k) = newKp

                a(k, q) = newKq
                a(q, k) = newKq

            Next

            a(p, p) =
                c * c * app -
                2.0 * s * c * apq +
                s * s * aqq

            a(q, q) =
                s * s * app +
                2.0 * s * c * apq +
                c * c * aqq

            a(p, q) = 0.0
            a(q, p) = 0.0

            For k As Integer = 0 To 2

                Dim vkp As Double = v(k, p)
                Dim vkq As Double = v(k, q)

                v(k, p) =
                    c * vkp -
                    s * vkq

                v(k, q) =
                    s * vkp +
                    c * vkq

            Next

        Next

        Dim values() As Double = {
            a(0, 0),
            a(1, 1),
            a(2, 2)
        }

        Dim order() As Integer = {
            0, 1, 2
        }

        'Sort ascending principal moments.
        For i As Integer = 0 To 1
            For j As Integer = i + 1 To 2

                If values(order(j)) <
                   values(order(i)) Then

                    Dim temp As Integer =
                        order(i)

                    order(i) =
                        order(j)

                    order(j) =
                        temp

                End If

            Next
        Next

        ReDim eigenValues(2)
        ReDim eigenVectors(2, 2)

        For newCol As Integer = 0 To 2

            Dim oldCol As Integer =
                order(newCol)

            eigenValues(newCol) =
                values(oldCol)

            Dim norm As Double =
                Math.Sqrt(
                    v(0, oldCol) * v(0, oldCol) +
                    v(1, oldCol) * v(1, oldCol) +
                    v(2, oldCol) * v(2, oldCol))

            If norm <= 0.0 Then
                norm = 1.0
            End If

            eigenVectors(0, newCol) =
                v(0, oldCol) / norm

            eigenVectors(1, newCol) =
                v(1, oldCol) / norm

            eigenVectors(2, newCol) =
                v(2, oldCol) / norm

        Next

    End Sub

    '==========================================================================
    ' MATERIAL / DENSITY
    '==========================================================================

    Private Function GetAssignedMaterialName(
        ByVal part As Part,
        ByVal body As Body) As String

        If part Is Nothing OrElse body Is Nothing Then
            Return ""
        End If

        Try

            Dim physicalMaterial As PhysicalMaterial =
                part.MaterialManager.
                    PhysicalMaterials.
                    AskMaterialOfObject(body)

            If physicalMaterial Is Nothing Then
                Return ""
            End If

            Try
                If Not String.IsNullOrWhiteSpace(
                    physicalMaterial.Name) Then

                    Return physicalMaterial.Name
                End If
            Catch
            End Try

            Return physicalMaterial.JournalIdentifier

        Catch
            Return ""
        End Try

    End Function

    Private Function GetBodyDensityKgM3(
        ByVal body As Body) As Double

        If body Is Nothing Then
            Return 0.0
        End If

        Try
            'NXOpen Body.Density is documented in kg/m^3.
            Return body.Density
        Catch
            Return 0.0
        End Try

    End Function

    Private Function GetFallbackDensityKgM3() As Double

        If fallbackDensityPrompted Then

            If Double.IsNaN(fallbackDensityKgM3) OrElse
               fallbackDensityKgM3 <= 0.0 Then

                Throw New OperationCanceledException(
                    "Valid fallback density was not provided.")
            End If

            Return fallbackDensityKgM3

        End If

        fallbackDensityPrompted = True

        Dim textValue As String =
            Interaction.InputBox(
                "One or more solid bodies do not have an assigned NX Physical Material." &
                Environment.NewLine &
                Environment.NewLine &
                "Enter fallback density in kg/m^3." &
                Environment.NewLine &
                "Example steel: 7850" &
                Environment.NewLine &
                Environment.NewLine &
                "This value is used only for the report; the NX model is not modified.",
                "Fallback Density",
                "7850")

        If String.IsNullOrWhiteSpace(textValue) Then

            fallbackDensityKgM3 = Double.NaN

            Throw New OperationCanceledException(
                "Fallback density entry cancelled.")

        End If

        Dim density As Double

        If Not Double.TryParse(
            textValue,
            NumberStyles.Float,
            CultureInfo.CurrentCulture,
            density) Then

            If Not Double.TryParse(
                textValue,
                NumberStyles.Float,
                CultureInfo.InvariantCulture,
                density) Then

                fallbackDensityKgM3 = Double.NaN

                Throw New Exception(
                    "Invalid fallback density: " &
                    textValue)

            End If

        End If

        If density <= 0.0 Then

            fallbackDensityKgM3 = Double.NaN

            Throw New Exception(
                "Fallback density must be greater than zero.")

        End If

        fallbackDensityKgM3 = density

        LogWarning(
            "Fallback density assumption enabled: " &
            density.ToString("0.###") &
            " kg/m^3.")

        Return fallbackDensityKgM3

    End Function

    '==========================================================================
    ' PART NUMBER / ATTRIBUTES
    '==========================================================================

    Private Function FindPartNumber(
        ByVal component As Component,
        ByVal part As Part) As String

        If component IsNot Nothing Then

            Try

                Dim value As String =
                    FindAttributeValue(
                        component.GetInstanceUserAttributes())

                If value <> "" Then
                    Return value
                End If

            Catch
            End Try

            Try

                Dim value As String =
                    FindAttributeValue(
                        component.GetUserAttributes())

                If value <> "" Then
                    Return value
                End If

            Catch
            End Try

        End If

        If part IsNot Nothing Then

            Try

                Dim value As String =
                    FindAttributeValue(
                        part.GetUserAttributes())

                If value <> "" Then
                    Return value
                End If

            Catch
            End Try

        End If

        Return ""

    End Function

    Private Function FindAttributeValue(
        ByVal attributes() As NXObject.AttributeInformation) As String

        If attributes Is Nothing Then Return ""

        For Each info As NXObject.AttributeInformation In attributes

            If info.Unset Then Continue For

            For Each aliasName As String In PART_NUMBER_ALIASES

                If String.Equals(
                    info.Title.Trim(),
                    aliasName,
                    StringComparison.OrdinalIgnoreCase) Then

                    Return AttributeValueAsString(info)

                End If

            Next

        Next

        Return ""

    End Function

    Private Function AttributeValueAsString(
        ByVal info As NXObject.AttributeInformation) As String

        Try

            Select Case info.Type

                Case NXObject.AttributeType.String
                    Return info.StringValue

                Case NXObject.AttributeType.Integer
                    Return info.IntegerValue.ToString(
                        CultureInfo.InvariantCulture)

                Case NXObject.AttributeType.Real
                    Return info.RealValue.ToString(
                        "0.############",
                        CultureInfo.InvariantCulture)

                Case Else
                    Return info.StringValue

            End Select

        Catch
            Return ""
        End Try

    End Function

    '==========================================================================
    ' COMPONENT / DISPLAY HELPERS
    '==========================================================================

    Private Function SafeIsSolid(
        ByVal body As Body) As Boolean

        If body Is Nothing Then Return False

        Try
            Return body.IsSolidBody
        Catch
            Return False
        End Try

    End Function

    Private Function SafeIsBlanked(
        ByVal displayObject As DisplayableObject) As Boolean

        If displayObject Is Nothing Then Return False

        Try
            Return displayObject.IsBlanked
        Catch
            Return False
        End Try

    End Function

    Private Function SafeComponentName(
        ByVal component As Component) As String

        If component Is Nothing Then
            Return "<WORK PART>"
        End If

        Try
            If Not String.IsNullOrWhiteSpace(
                component.DisplayName) Then

                Return component.DisplayName
            End If
        Catch
        End Try

        Try
            If Not String.IsNullOrWhiteSpace(
                component.Name) Then

                Return component.Name
            End If
        Catch
        End Try

        Return "Component Tag " &
               component.Tag.ToString()

    End Function

    Private Function ComponentPath(
        ByVal component As Component) As String

        If component Is Nothing Then
            Return modelPart.Leaf
        End If

        Dim names As New List(Of String)

        Dim current As Component =
            component

        While current IsNot Nothing

            names.Add(
                SafeComponentName(current))

            Try
                current = current.Parent
            Catch
                current = Nothing
            End Try

        End While

        names.Reverse()

        Return String.Join("/", names.ToArray())

    End Function

    '==========================================================================
    ' MATERIAL LABEL AGGREGATION
    '==========================================================================

    Private Function AggregateMaterialLabel(
        ByVal members As List(Of MassResult)) As String

        If members Is Nothing OrElse members.Count = 0 Then
            Return ""
        End If

        Dim first As String = Nothing

        For Each member As MassResult In members

            If member Is Nothing Then Continue For

            Dim current As String =
                If(
                    String.IsNullOrWhiteSpace(member.MaterialName),
                    "<UNSPECIFIED>",
                    member.MaterialName)

            If first Is Nothing Then
                first = current

            ElseIf Not String.Equals(
                first,
                current,
                StringComparison.OrdinalIgnoreCase) Then

                Return "MIXED"

            End If

        Next

        Return If(first, "")

    End Function

    Private Function AppendNote(
        ByVal existing As String,
        ByVal newNote As String) As String

        If String.IsNullOrWhiteSpace(existing) Then
            Return newNote
        End If

        If String.IsNullOrWhiteSpace(newNote) Then
            Return existing
        End If

        Return existing & " " & newNote

    End Function

    '==========================================================================
    ' LISTING WINDOW REPORT
    '==========================================================================

    Private Sub PrintHeader()

        lw.WriteLine("")
        lw.WriteLine("======================================================================")
        lw.WriteLine(" NX 2306 - MASS PROPERTIES & INERTIA REPORT")
        lw.WriteLine("======================================================================")
        lw.WriteLine("")
        lw.WriteLine("Report units:")
        lw.WriteLine("  Mass                 = kg")
        lw.WriteLine("  Gravitational weight = N (g = " &
                     STANDARD_GRAVITY.ToString("0.#####") &
                     " m/s^2)")
        lw.WriteLine("  CG                   = mm")
        lw.WriteLine("  Volume               = mm^3")
        lw.WriteLine("  Density              = kg/m^3 and kg/mm^3")
        lw.WriteLine("  Inertia              = kg*mm^2")
        lw.WriteLine("")
        lw.WriteLine("Reference:")
        lw.WriteLine("  CG and inertia axes = displayed part ABSOLUTE XYZ.")
        lw.WriteLine("  Origin inertia      = about absolute origin (0,0,0).")
        lw.WriteLine("  CG inertia          = about combined CG, axes parallel to absolute XYZ.")
        lw.WriteLine("")
        lw.WriteLine("Products of inertia in this report:")
        lw.WriteLine("  Pxy = +Integral(x*y dm), Pxz = +Integral(x*z dm), Pyz = +Integral(y*z dm)")
        lw.WriteLine("  Inertia tensor off-diagonal terms are -Pxy, -Pxz, -Pyz.")
        lw.WriteLine("")

    End Sub

    Private Sub PrintAllResults()

        lw.WriteLine("")
        lw.WriteLine("======================================================================")
        lw.WriteLine(" RESULTS")
        lw.WriteLine("======================================================================")

        For Each result As MassResult In reportRows

            PrintResult(result)

        Next

    End Sub

    Private Sub PrintResult(
        ByVal r As MassResult)

        lw.WriteLine("")
        lw.WriteLine("----------------------------------------------------------------------")
        lw.WriteLine(r.RowType & " : " & r.ComponentName)
        lw.WriteLine("----------------------------------------------------------------------")

        lw.WriteLine("Part name       : " & r.PartName)
        lw.WriteLine("Component       : " & r.ComponentName)
        lw.WriteLine("Component path  : " & r.ComponentPath)
        lw.WriteLine("Part number     : " & r.PartNumber)
        lw.WriteLine("Hidden/blanked  : " & r.Hidden.ToString())

        lw.WriteLine("Material        : " & r.MaterialName)
        lw.WriteLine("Density source  : " & r.DensitySource)
        lw.WriteLine("Assumed density : " & r.UsedAssumedDensity.ToString())

        lw.WriteLine(
            "Mass             : " &
            F(r.MassKg) &
            " kg")

        lw.WriteLine(
            "Weight (gravity) : " &
            F(r.WeightN) &
            " N")

        lw.WriteLine(
            "Volume           : " &
            F(r.VolumeMm3) &
            " mm^3")

        lw.WriteLine(
            "Density          : " &
            F(r.DensityKgM3) &
            " kg/m^3")

        lw.WriteLine(
            "Density          : " &
            F(r.DensityKgMm3) &
            " kg/mm^3")

        lw.WriteLine("")
        lw.WriteLine("CG relative to absolute origin [mm]:")

        lw.WriteLine(
            "  X = " & F(r.CgX) &
            "   Y = " & F(r.CgY) &
            "   Z = " & F(r.CgZ))

        lw.WriteLine("")
        lw.WriteLine("Inertia about CG, axes parallel to absolute XYZ [kg*mm^2]:")

        lw.WriteLine(
            "  Ixx = " & F(r.IxxCg) &
            "   Iyy = " & F(r.IyyCg) &
            "   Izz = " & F(r.IzzCg))

        lw.WriteLine(
            "  Pxy = " & F(r.PxyCg) &
            "   Pxz = " & F(r.PxzCg) &
            "   Pyz = " & F(r.PyzCg))

        lw.WriteLine("")
        lw.WriteLine("Inertia about absolute origin [kg*mm^2]:")

        lw.WriteLine(
            "  Ixx = " & F(r.IxxOrigin) &
            "   Iyy = " & F(r.IyyOrigin) &
            "   Izz = " & F(r.IzzOrigin))

        lw.WriteLine(
            "  Pxy = " & F(r.PxyOrigin) &
            "   Pxz = " & F(r.PxzOrigin) &
            "   Pyz = " & F(r.PyzOrigin))

        lw.WriteLine("")
        lw.WriteLine("Principal moments about CG [kg*mm^2]:")

        lw.WriteLine(
            "  I1 = " & F(r.Principal1) &
            "   I2 = " & F(r.Principal2) &
            "   I3 = " & F(r.Principal3))

        lw.WriteLine("Principal axis directions relative to absolute XYZ:")

        lw.WriteLine(
            "  Axis 1 = (" &
            F(r.Axis1X) & ", " &
            F(r.Axis1Y) & ", " &
            F(r.Axis1Z) & ")")

        lw.WriteLine(
            "  Axis 2 = (" &
            F(r.Axis2X) & ", " &
            F(r.Axis2Y) & ", " &
            F(r.Axis2Z) & ")")

        lw.WriteLine(
            "  Axis 3 = (" &
            F(r.Axis3X) & ", " &
            F(r.Axis3Y) & ", " &
            F(r.Axis3Z) & ")")

        If Not String.IsNullOrWhiteSpace(r.Notes) Then
            lw.WriteLine("Notes            : " & r.Notes)
        End If

    End Sub

    Private Sub PrintRunSummary()

        lw.WriteLine("")
        lw.WriteLine("======================================================================")
        lw.WriteLine(" RUN SUMMARY")
        lw.WriteLine("======================================================================")
        lw.WriteLine("Rows generated             : " & reportRows.Count.ToString())
        lw.WriteLine("Suppressed components skip : " & suppressedComponentCount.ToString())
        lw.WriteLine("Hidden components seen     : " & hiddenComponentCount.ToString())
        lw.WriteLine("Unresolved occurrence body : " & unresolvedBodyCount.ToString())
        lw.WriteLine("Non-solid geometry skipped : " & skippedNonSolidCount.ToString())
        lw.WriteLine("Warnings                   : " & warningCount.ToString())
        lw.WriteLine("Errors                     : " & errorCount.ToString())
        lw.WriteLine("CSV                         : " & csvPath)
        lw.WriteLine("======================================================================")
        lw.WriteLine("")

    End Sub

    '==========================================================================
    ' CSV
    '==========================================================================

    Private Sub ExportCsv()

        If String.IsNullOrWhiteSpace(csvFolder) Then
            Throw New Exception("CSV output folder is not defined.")
        End If

        Dim baseName As String =
            SafeFileName(
                Path.GetFileNameWithoutExtension(modelPart.Leaf))

        csvPath =
            Path.Combine(
                csvFolder,
                baseName &
                "_Mass_Inertia_Report_" &
                DateTime.Now.ToString("yyyyMMdd_HHmmss") &
                ".csv")

        Using writer As New StreamWriter(
            csvPath,
            False,
            New UTF8Encoding(True))

            writer.WriteLine(
                String.Join(
                    ",",
                    New String() {
                        "RowType",
                        "PartName",
                        "ComponentName",
                        "ComponentPath",
                        "PartNumber",
                        "Hidden",
                        "Material",
                        "DensitySource",
                        "AssumedDensity",
                        "Mass_kg",
                        "Weight_N",
                        "Volume_mm3",
                        "Density_kg_m3",
                        "Density_kg_mm3",
                        "CG_X_mm",
                        "CG_Y_mm",
                        "CG_Z_mm",
                        "Ixx_CG_kg_mm2",
                        "Iyy_CG_kg_mm2",
                        "Izz_CG_kg_mm2",
                        "Pxy_CG_PositiveConvention_kg_mm2",
                        "Pxz_CG_PositiveConvention_kg_mm2",
                        "Pyz_CG_PositiveConvention_kg_mm2",
                        "Ixx_Origin_kg_mm2",
                        "Iyy_Origin_kg_mm2",
                        "Izz_Origin_kg_mm2",
                        "Pxy_Origin_PositiveConvention_kg_mm2",
                        "Pxz_Origin_PositiveConvention_kg_mm2",
                        "Pyz_Origin_PositiveConvention_kg_mm2",
                        "Principal_I1_kg_mm2",
                        "Principal_I2_kg_mm2",
                        "Principal_I3_kg_mm2",
                        "Axis1_X",
                        "Axis1_Y",
                        "Axis1_Z",
                        "Axis2_X",
                        "Axis2_Y",
                        "Axis2_Z",
                        "Axis3_X",
                        "Axis3_Y",
                        "Axis3_Z",
                        "Notes"
                    }))

            For Each r As MassResult In reportRows

                Dim fields As New List(Of String)

                fields.Add(Csv(r.RowType))
                fields.Add(Csv(r.PartName))
                fields.Add(Csv(r.ComponentName))
                fields.Add(Csv(r.ComponentPath))
                fields.Add(Csv(r.PartNumber))
                fields.Add(Csv(r.Hidden.ToString()))
                fields.Add(Csv(r.MaterialName))
                fields.Add(Csv(r.DensitySource))
                fields.Add(Csv(r.UsedAssumedDensity.ToString()))

                fields.Add(N(r.MassKg))
                fields.Add(N(r.WeightN))
                fields.Add(N(r.VolumeMm3))
                fields.Add(N(r.DensityKgM3))
                fields.Add(N(r.DensityKgMm3))

                fields.Add(N(r.CgX))
                fields.Add(N(r.CgY))
                fields.Add(N(r.CgZ))

                fields.Add(N(r.IxxCg))
                fields.Add(N(r.IyyCg))
                fields.Add(N(r.IzzCg))

                fields.Add(N(r.PxyCg))
                fields.Add(N(r.PxzCg))
                fields.Add(N(r.PyzCg))

                fields.Add(N(r.IxxOrigin))
                fields.Add(N(r.IyyOrigin))
                fields.Add(N(r.IzzOrigin))

                fields.Add(N(r.PxyOrigin))
                fields.Add(N(r.PxzOrigin))
                fields.Add(N(r.PyzOrigin))

                fields.Add(N(r.Principal1))
                fields.Add(N(r.Principal2))
                fields.Add(N(r.Principal3))

                fields.Add(N(r.Axis1X))
                fields.Add(N(r.Axis1Y))
                fields.Add(N(r.Axis1Z))

                fields.Add(N(r.Axis2X))
                fields.Add(N(r.Axis2Y))
                fields.Add(N(r.Axis2Z))

                fields.Add(N(r.Axis3X))
                fields.Add(N(r.Axis3Y))
                fields.Add(N(r.Axis3Z))

                fields.Add(Csv(r.Notes))

                writer.WriteLine(
                    String.Join(
                        ",",
                        fields.ToArray()))

            Next

        End Using

        LogInfo("CSV exported: " & csvPath)

    End Sub

    Private Function Csv(
        ByVal text As String) As String

        If text Is Nothing Then text = ""

        Return """" &
               text.Replace("""", """""") &
               """"

    End Function

    Private Function N(
        ByVal value As Double) As String

        Return value.ToString(
            "0.###############E+0",
            CultureInfo.InvariantCulture)

    End Function

    Private Function F(
        ByVal value As Double) As String

        Return value.ToString(
            "0.######",
            CultureInfo.InvariantCulture)

    End Function

    Private Function SafeFileName(
        ByVal text As String) As String

        If String.IsNullOrWhiteSpace(text) Then
            Return "NX_Model"
        End If

        Dim result As String = text

        For Each c As Char In Path.GetInvalidFileNameChars()
            result = result.Replace(c, "_"c)
        Next

        Return result

    End Function

    '==========================================================================
    ' LOGGING
    '==========================================================================

    Private Sub LogInfo(
        ByVal message As String)

        lw.WriteLine(message)

    End Sub

    Private Sub LogWarning(
        ByVal message As String)

        warningCount += 1

        lw.WriteLine(
            "[WARNING] " &
            message)

    End Sub

    Private Sub LogError(
        ByVal message As String)

        errorCount += 1

        lw.WriteLine(
            "[ERROR] " &
            message)

    End Sub

    '==========================================================================
    ' UNLOAD
    '==========================================================================

    Public Function GetUnloadOption(
        ByVal dummy As String) As Integer

        Return Session.LibraryUnloadOption.Immediately

    End Function

End Module
