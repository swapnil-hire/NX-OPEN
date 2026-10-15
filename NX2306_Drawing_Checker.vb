' NX 2306 native drawing inspection - staged implementation 0.1, 2026-10-15
' Full source. Tools > Journal > Play. Windows NX + desktop Excel for XLSX.
' Read README.md and COVERAGE.md. NOT a release approval/certification tool.
' Reviewed against SC2306 and older Siemens-authored documentation mirrors.
' Not compiled or executed against installed NX 2306 DLLs in this environment.
' Does not save, close, update, recolour or edit source drawings.
' Reference is opened in background and remains open. No automatic part close.
' Only transient Highlight/Unhighlight; no category colours are written to NX.
' Exact, engineering and configured-company-rule modes have bounded coverage.
Option Strict Off ' Excel COM late binding only; NX calls below are typed.
Option Explicit On
Imports System
Imports System.IO
Imports System.Text
Imports System.Collections.Generic
Imports System.Globalization
Imports System.Text.RegularExpressions
Imports System.Security.Cryptography
Imports System.Reflection
Imports System.Runtime.InteropServices
Imports NXOpen
Imports NXOpen.UF
Imports NXOpen.Utilities
Imports NXOpen.Annotations
Imports NXOpen.Drawings
Imports NXOpen.Gateway
Imports F = System.Windows.Forms
Imports D = System.Drawing

Module DrawingChecker
    Friend Session As NXOpen.Session
    Friend UF As UFSession
    Friend Current As Part
    Friend Diagnostics As New List(Of String)()
    Friend ActiveHighlight As DisplayableObject
    Friend Const Version As String = "0.1 - staged native inspection"
    Friend Const IdAttribute As String = "CHECK_ID"

    Public Sub Main()
        Session = NXOpen.Session.GetSession() : UF = UFSession.GetUFSession()
        Session.ListingWindow.Open() : Diagnostics.Clear()
        Try
            EngineTests.Run()
            Log("Pure comparison-engine startup self-checks passed.")
            Current = Session.Parts.Work
            If Current Is Nothing OrElse Session.Parts.Display Is Nothing Then Throw New Exception("Open a native NX drawing first.")
            If Current.Tag <> Session.Parts.Display.Tag Then Throw New Exception("Make the displayed drawing the work part.")
            If Current.DrawingSheets.ToArray().Length = 0 Then Throw New Exception("The work part has no drawing sheets.")
            Log("NX drawing checker " & Version)
            Log("NXOpen assembly: " & GetType(NXObject).Assembly.GetName().Version.ToString())
            Log("Current: " & Current.FullPath)
            Using form As New CheckerForm(Current)
                form.ShowDialog()
            End Using
        Catch ex As Exception
            Log("ERROR: " & ex.ToString())
            F.MessageBox.Show(ex.Message, "NX drawing checker", F.MessageBoxButtons.OK, F.MessageBoxIcon.Error)
        Finally
            ClearHighlight()
        End Try
    End Sub
    Friend Sub Log(text As String)
        Diagnostics.Add(DateTime.Now.ToString("s") & " " & text)
        Session.ListingWindow.WriteLine(text)
    End Sub
    Friend Sub ClearHighlight()
        If ActiveHighlight Is Nothing Then Return
        Try
            ActiveHighlight.Unhighlight()
        Catch ex As Exception
            Log("WARNING clearing transient highlight: " & ex.Message)
        Finally
            ActiveHighlight = Nothing
        End Try
    End Sub
    Public Function GetUnloadOption(dummy As String) As Integer
        Return CInt(NXOpen.Session.LibraryUnloadOption.Immediately)
    End Function
End Module

' -------------------- Structured model / pure comparison engine --------------------
Public Class Item
    Public Key As String = "", Category As String = "", Kind As String = "", Name As String = ""
    Public CheckId As String = "", JournalId As String = "", Sheet As String = "", ViewName As String = ""
    Public RawText As String = "", Unit As String = "", NumericUnit As String = "raw API value", State As String = "OK", Problem As String = ""
    Public X As Double, Y As Double, Z As Double
    Public HasLocation As Boolean
    Public Fields As New SortedDictionary(Of String, String)(StringComparer.Ordinal)
    Public Numbers As New SortedDictionary(Of String, Double)(StringComparer.Ordinal)
    Public Datums As New List(Of String)()
    Public Native As NXObject
    Public Sub Put(name As String, value As Object)
        Fields(name) = Convert.ToString(value, CultureInfo.InvariantCulture)
    End Sub
    Public Function Identity() As String
        If CheckId <> "" Then Return Category & ":CHECK_ID=" & CheckId
        If JournalId <> "" Then Return Category & ":JOURNAL_ID=" & JournalId
        Return Category & ":SESSION_TAG=" & Key
    End Function
    Public Function Signature() As String
        Dim b As New StringBuilder(Category & "|" & Kind & "|" & Unit & "|" & RawText)
        For Each pair As KeyValuePair(Of String, String) In Fields
            b.Append("|" & pair.Key & "=" & pair.Value)
        Next
        For Each pair As KeyValuePair(Of String, Double) In Numbers
            b.Append("|" & pair.Key & "=" & pair.Value.ToString("R", CultureInfo.InvariantCulture))
        Next
        Return b.ToString()
    End Function
    Public Function Describe() As String
        Dim b As New StringBuilder(Kind & " " & Name & " " & RawText)
        For Each p As KeyValuePair(Of String, Double) In Numbers
            b.Append("; " & p.Key & "=" & p.Value.ToString("G12", CultureInfo.InvariantCulture))
        Next
        For Each p As KeyValuePair(Of String, String) In Fields
            b.Append("; " & p.Key & "=" & p.Value)
        Next
        Return b.ToString()
    End Function
End Class
Public Class Finding
    Public ID As String, Category As String, Rule As String, Status As String
    Public Reference As String, CurrentValue As String, Message As String, MatchMethod As String
    Public RefItem As Item, CurItem As Item
End Class
Public Class Settings
    Public Mode As String = "Engineering"
    Public NumericEpsilon As Double = 0.000001 ' computational equality; never an allowed design tolerance
    Public LayoutMm As Double = 1.0
    Public TrustJournalIds As Boolean
    Public Categories As New HashSet(Of String)(StringComparer.Ordinal)
    Public CurrentSheets As New HashSet(Of String)(StringComparer.Ordinal)
    Public ReferenceSheets As New HashSet(Of String)(StringComparer.Ordinal)
    Public CurrentViews As New HashSet(Of String)(StringComparer.Ordinal)
    Public ReferenceViews As New HashSet(Of String)(StringComparer.Ordinal)
    Public CurrentAllSheets As Boolean = True, ReferenceAllSheets As Boolean = True
    Public CurrentAllViews As Boolean = True, ReferenceAllViews As Boolean = True
    Public ScopeSheetNotes As Boolean = True
    Public ProfilePath As String = "", DictionaryPath As String = ""
End Class
Public Class Snapshot
    Public Label As String, PartPath As String, PartUnit As String
    Public Items As New List(Of Item)()
    Public Issues As New List(Of Finding)()
End Class
Public Class Inspector
    Public Shared Function Normalize(text As String) As String
        ' Whitespace only; preserve case, technical wording, punctuation and symbol tokens.
        Return Regex.Replace(If(text, "").Replace(vbCrLf, vbLf).Replace(vbCr, vbLf), "\s+", " ").Trim()
    End Function
    Public Shared Function Make(status As String, category As String, rule As String, r As Item, c As Item, oldValue As String, newValue As String, message As String, Optional method As String = "") As Finding
        Dim seed As String = category & "|" & rule & "|" & If(r Is Nothing, "-", r.Identity()) & "|" & If(c Is Nothing, "-", c.Identity())
        Dim digest() As Byte
        Using hash As SHA256 = SHA256.Create()
            digest = hash.ComputeHash(Encoding.UTF8.GetBytes(seed))
        End Using
        Dim id As String = BitConverter.ToString(digest).Replace("-", "").Substring(0, 20)
        Return New Finding With {.ID = id, .Status = status, .Category = category, .Rule = rule, .RefItem = r, .CurItem = c, .Reference = oldValue, .CurrentValue = newValue, .Message = message, .MatchMethod = method}
    End Function
    Public Shared Function Compare(reference As Snapshot, current As Snapshot, options As Settings, Optional progress As Action(Of String) = Nothing, Optional cancelled As Func(Of Boolean) = Nothing) As List(Of Finding)
        Dim answer As New List(Of Finding)()
        answer.AddRange(reference.Issues) : answer.AddRange(current.Issues)
        Dim rr As New List(Of Item)(reference.Items), cc As New List(Of Item)(current.Items)
        ' Each stage uses a complete candidate graph, then accepts only mutual unique edges.
        ' Never use order, nearest position or a single JournalIdentifier as universal identity.
        PairStage(rr, cc, answer, options, progress, cancelled, "CHECK_ID", Function(a, b) a.Category = b.Category AndAlso a.CheckId <> "" AndAlso a.CheckId = b.CheckId)
        If options.TrustJournalIds Then PairStage(rr, cc, answer, options, progress, cancelled, "trusted related-revision journal ID", Function(a, b) a.Category = b.Category AndAlso a.Kind = b.Kind AndAlso a.JournalId <> "" AndAlso a.JournalId = b.JournalId)
        PairStage(rr, cc, answer, options, progress, cancelled, "unique object name (review)", Function(a, b) a.Category = b.Category AndAlso a.Kind = b.Kind AndAlso a.Name <> "" AndAlso a.Name = b.Name)
        PairStage(rr, cc, answer, options, progress, cancelled, "unique content (review)", Function(a, b) a.Signature() = b.Signature())
        For Each r As Item In rr
            Dim candidates As New List(Of Item)()
            For Each c As Item In cc
                If c.Category = r.Category AndAlso c.Kind = r.Kind Then candidates.Add(c)
            Next
            If candidates.Count = 0 Then
                answer.Add(Make("MISSING", r.Category, "Unmatched reference", r, Nothing, r.Describe(), "", "No remaining same-kind candidate in selected scope; review extraction/coverage errors before confirming absence."))
            Else
                answer.Add(Make("WARNING", r.Category, "Ambiguous correspondence", r, Nothing, r.Describe(), candidates.Count.ToString() & " possible current objects", "No automatic match. Assign common unique CHECK_ID values or review manually. Position does not establish engineering identity."))
            End If
        Next
        For Each c As Item In cc
            Dim possible As Boolean = False
            For Each r As Item In rr
                If r.Category = c.Category AndAlso r.Kind = c.Kind Then possible = True
            Next
            answer.Add(Make(If(possible, "WARNING", "EXTRA"), c.Category, "Unmatched current", Nothing, c, "", c.Describe(), If(possible, "Unresolved correspondence; not classified as an addition.", "No remaining same-kind reference candidate in selected scope.")))
        Next
        If current.Items.Count = 0 OrElse reference.Items.Count = 0 Then answer.Add(Make("UNVERIFIED", "Coverage", "Empty inventory", Nothing, Nothing, reference.Items.Count.ToString(), current.Items.Count.ToString(), "Empty inventories cannot establish conformity."))
        Return answer
    End Function
    Private Shared Sub PairStage(rr As List(Of Item), cc As List(Of Item), answer As List(Of Finding), options As Settings, progress As Action(Of String), cancelled As Func(Of Boolean), method As String, predicate As Func(Of Item, Item, Boolean))
        Dim pairs As New Dictionary(Of Item, Item)()
        Dim scanned As Integer = 0
        For Each r As Item In rr
            scanned += 1
            If scanned Mod 50 = 0 Then
                If progress IsNot Nothing Then progress("Matching: " & method & " " & scanned.ToString() & "/" & rr.Count.ToString())
                If cancelled IsNot Nothing AndAlso cancelled() Then Throw New OperationCanceledException()
            End If
            Dim match As Item = Nothing, count As Integer = 0
            For Each c As Item In cc
                If predicate(r, c) Then
                    count += 1 : match = c
                End If
            Next
            If count <> 1 Then Continue For
            Dim reverse As Integer = 0
            For Each other As Item In rr
                If predicate(other, match) Then reverse += 1
            Next
            If reverse = 1 Then pairs.Add(r, match)
        Next
        For Each pair As KeyValuePair(Of Item, Item) In pairs
            ComparePair(pair.Key, pair.Value, answer, options, method)
            rr.Remove(pair.Key) : cc.Remove(pair.Value)
        Next
    End Sub
    Private Shared Sub ComparePair(r As Item, c As Item, answer As List(Of Finding), options As Settings, method As String)
        If r.Kind <> c.Kind Then answer.Add(Make("FAIL", r.Category, "NX annotation type", r, c, r.Kind, c.Kind, "Objects share the selected identity but use different NX object types.", method))
        If method.Contains("review") Then answer.Add(Make("WARNING", r.Category, "Correspondence review", r, c, r.Identity(), c.Identity(), "Candidate paired by " & method & "; confirm that the same feature/requirement is represented.", method))
        If r.State <> "OK" OrElse c.State <> "OK" Then answer.Add(Make("UNVERIFIED", r.Category, "Incomplete record", r, c, r.Problem, c.Problem, "Only successfully extracted fields are compared.", method))
        Dim fields As New SortedSet(Of String)(r.Fields.Keys, StringComparer.Ordinal)
        fields.UnionWith(c.Fields.Keys)
        For Each field As String In fields
            Dim a As String = Nothing, b As String = Nothing
            Dim hasA As Boolean = r.Fields.TryGetValue(field, a), hasB As Boolean = c.Fields.TryGetValue(field, b)
            Dim severity As String = If(hasA AndAlso hasB, If(a = b, "PASS", "FAIL"), "UNVERIFIED")
            If field.StartsWith("Format.", StringComparison.Ordinal) AndAlso severity = "FAIL" AndAlso options.Mode <> "Exact" Then severity = "WARNING"
            answer.Add(Make(severity, r.Category, field, r, c, If(a, "<unavailable>"), If(b, "<unavailable>"), "Structured field comparison.", method))
        Next
        Dim numbers As New SortedSet(Of String)(r.Numbers.Keys, StringComparer.Ordinal)
        numbers.UnionWith(c.Numbers.Keys)
        For Each field As String In numbers
            Dim a As Double = 0, b As Double = 0
            Dim ok As Boolean = r.Numbers.TryGetValue(field, a) AndAlso c.Numbers.TryGetValue(field, b)
            Dim status As String = "UNVERIFIED"
            If ok AndAlso r.Unit = c.Unit AndAlso r.NumericUnit = c.NumericUnit AndAlso Not Double.IsNaN(a) AndAlso Not Double.IsNaN(b) AndAlso Not Double.IsInfinity(a) AndAlso Not Double.IsInfinity(b) Then status = If(Math.Abs(a - b) <= options.NumericEpsilon, "PASS", "FAIL")
            answer.Add(Make(status, r.Category, field, r, c, a.ToString("G15", CultureInfo.InvariantCulture) & " [" & r.NumericUnit & "; part=" & r.Unit & "]", b.ToString("G15", CultureInfo.InvariantCulture) & " [" & c.NumericUnit & "; part=" & c.Unit & "]", "Native API values. Mixed part units are not converted; numeric epsilon is computational equality, not permitted tolerance growth.", method))
        Next
        If r.RawText <> "" OrElse c.RawText <> "" Then
            Dim exact As Boolean = r.RawText = c.RawText
            Dim normalized As Boolean = Normalize(r.RawText) = Normalize(c.RawText)
            Dim status As String = If(exact, "PASS", If(normalized AndAlso options.Mode <> "Exact", "WARNING", "FAIL"))
            answer.Add(Make(status, r.Category, "Displayed text", r, c, r.RawText, c.RawText, "Exact=" & exact.ToString() & "; whitespace-normalized=" & normalized.ToString() & ". No synonym, symbol or engineering-value substitution is applied.", method))
        End If
        If r.Sheet <> c.Sheet OrElse r.ViewName <> c.ViewName Then answer.Add(Make("WARNING", r.Category, "Sheet/view membership", r, c, r.Sheet & "/" & r.ViewName, c.Sheet & "/" & c.ViewName, "Layout/context changed; no automatic engineering conclusion.", method))
        If r.HasLocation AndAlso c.HasLocation AndAlso r.Unit = c.Unit Then
            Dim shift As Double = Math.Sqrt((r.X - c.X) ^ 2 + (r.Y - c.Y) ^ 2 + (r.Z - c.Z) ^ 2)
            Dim mm As Double = If(r.Unit = "Inches", shift * 25.4, shift)
            If mm > options.LayoutMm Then answer.Add(Make("WARNING", r.Category, "Location movement", r, c, Location(r), Location(c), "Layout movement " & mm.ToString("G8", CultureInfo.InvariantCulture) & " mm; attachment equivalence is not established.", method))
        End If
    End Sub
    Public Shared Function Location(i As Item) As String
        If i Is Nothing OrElse Not i.HasLocation Then Return "unresolved"
        Return String.Format(CultureInfo.InvariantCulture, "({0:G10}, {1:G10}, {2:G10}) {3}", i.X, i.Y, i.Z, i.Unit)
    End Function
    Public Shared Function Overall(rows As List(Of Finding), extractionOnly As Boolean) As String
        If extractionOnly Then Return "EXTRACTED - NOT INSPECTED"
        Dim failure As Boolean = False, incomplete As Boolean = False, warning As Boolean = False
        For Each row As Finding In rows
            If row.Status = "ERROR" OrElse row.Status = "UNVERIFIED" Then incomplete = True
            If row.Status = "FAIL" OrElse row.Status = "MISSING" Then failure = True
            If row.Status = "WARNING" OrElse row.Status = "EXTRA" Then warning = True
        Next
        If failure AndAlso incomplete Then Return "FAIL / INCOMPLETE"
        If failure Then Return "FAIL"
        If incomplete OrElse rows.Count = 0 Then Return "INCOMPLETE - NO OVERALL PASS"
        If warning Then Return "REVIEW REQUIRED"
        Return "PASS - SELECTED IMPLEMENTED CHECKS ONLY"
    End Function
End Class

' -------------------- NX extraction (no Commit / save / update) --------------------
Public Class Extractor
    Private Owner As Part, Options As Settings, Data As Snapshot
    Private ViewSheets As New Dictionary(Of Tag, String)()
    Private ViewNames As New Dictionary(Of Tag, String)()
    Private NativeViewNames As New Dictionary(Of Tag, String)()
    Private Progress As Action(Of String)
    Private Cancelled As Func(Of Boolean)
    Private IsReference As Boolean
    Public Sub New(p As Part, config As Settings, label As String, progressCallback As Action(Of String), cancelCallback As Func(Of Boolean), referenceSide As Boolean)
        Owner = p : Options = config : Progress = progressCallback : Cancelled = cancelCallback : IsReference = referenceSide
        Data = New Snapshot With {.Label = label, .PartPath = p.FullPath, .PartUnit = p.PartUnits.ToString()}
    End Sub
    Private Sub Problem(status As String, rule As String, i As Item, message As String)
        If i IsNot Nothing Then
            If status = "ERROR" OrElse i.State = "OK" OrElse (status = "UNVERIFIED" AndAlso i.State <> "ERROR") Then i.State = status
            i.Problem &= rule & ": " & message & "; "
        End If
        Data.Issues.Add(Inspector.Make(status, If(i Is Nothing, "Coverage", i.Category), Data.Label & ": " & rule, If(IsReference, i, Nothing), If(IsReference, Nothing, i), "", "", message))
        DrawingChecker.Log(Data.Label & " " & status & " " & rule & ": " & message)
    End Sub
    Public Function Run() As Snapshot
        If Not Owner.IsFullyLoaded Then Problem("UNVERIFIED", "Part load", Nothing, "Part is not fully loaded. Load all required drafting/model dependencies and rerun.")
        For Each sheet As DrawingSheet In Owner.DrawingSheets.ToArray()
            If Cancelled() Then Throw New OperationCanceledException()
            Progress(Data.Label & " sheet: " & sheet.Name)
            Try
                ViewSheets(sheet.View.Tag) = sheet.Name : ViewNames(sheet.View.Tag) = "<sheet>" : NativeViewNames(sheet.View.Tag) = sheet.View.Name
                For Each view As DraftingView In sheet.GetDraftingViews()
                    ViewSheets(view.Tag) = sheet.Name : ViewNames(view.Tag) = view.Name : NativeViewNames(view.Tag) = view.Name
                    If Options.Categories.Contains("Views") Then ExtractView(view, sheet)
                Next
                If Options.Categories.Contains("Drawing standards") Then ExtractSheet(sheet)
            Catch ex As Exception
                Problem("ERROR", "Sheet traversal " & sheet.Name, Nothing, ex.Message)
            End Try
        Next
        Dim tag As Tag = Tag.Null, count As Integer = 0
        Do
            tag = DrawingChecker.UF.Obj.CycleAll(Owner.Tag, tag)
            If tag = Tag.Null Then Exit Do
            count += 1
            If count Mod 100 = 0 Then
                Progress(Data.Label & " scanned " & count.ToString() & " native objects")
                If Cancelled() Then Throw New OperationCanceledException()
            End If
            Dim nativeType As Integer = 0, subtype As Integer = 0
            Try
                DrawingChecker.UF.Obj.AskTypeAndSubtype(tag, nativeType, subtype)
                Dim obj As NXObject = TryCast(NXObjectManager.Get(tag), NXObject)
                Dim annotation As Annotation = TryCast(obj, Annotation)
                If annotation IsNot Nothing Then
                    ExtractAnnotation(annotation)
                ElseIf (nativeType = UFConstants.UF_drafting_entity_type OrElse nativeType = UFConstants.UF_dimension_type) AndAlso Not (TypeOf obj Is NXOpen.View OrElse TypeOf obj Is DrawingSheet) Then
                    Problem("UNVERIFIED", "Unwrapped drafting object " & tag.ToString(), Nothing, "Native type/subtype " & nativeType.ToString() & "/" & subtype.ToString() & " is not represented by a supported Annotation wrapper. Could be a table, symbol or other drafting entity.")
                End If
            Catch ex As Exception
                Problem("ERROR", "Native object " & tag.ToString(), Nothing, "Type/subtype=" & nativeType.ToString() & "/" & subtype.ToString() & "; " & ex.Message)
            End Try
        Loop
        If Options.Categories.Contains("Title block") Then ExtractAttributes()
        ValidateDatums()
        DuplicateIds()
        AddCoverage()
        Progress(Data.Label & " extracted " & Data.Items.Count.ToString() & " records")
        Return Data
    End Function
    Private Function BaseItem(obj As NXObject, category As String) As Item
        Dim i As New Item With {.Native = obj, .Category = category, .Kind = obj.GetType().Name, .Key = obj.Tag.ToString(), .JournalId = obj.JournalIdentifier, .Name = obj.Name, .Unit = Data.PartUnit}
        If obj.HasUserAttribute(DrawingChecker.IdAttribute, NXObject.AttributeType.String, -1) Then i.CheckId = obj.GetStringUserAttribute(DrawingChecker.IdAttribute, -1)
        Return i
    End Function
    Private Function InScope(i As Item) As Boolean
        Dim allSheets As Boolean = If(IsReference, Options.ReferenceAllSheets, Options.CurrentAllSheets)
        Dim sheets As HashSet(Of String) = If(IsReference, Options.ReferenceSheets, Options.CurrentSheets)
        Dim allViews As Boolean = If(IsReference, Options.ReferenceAllViews, Options.CurrentAllViews)
        Dim views As HashSet(Of String) = If(IsReference, Options.ReferenceViews, Options.CurrentViews)
        ' Unresolved membership stays visible as incomplete; it is never silently dropped.
        If i.Sheet = "" Then Return True
        Dim sheetOK As Boolean = allSheets
        For Each name As String In i.Sheet.Split("|"c)
            If sheets.Contains(name) Then sheetOK = True
        Next
        If Not sheetOK Then Return False
        If allViews OrElse i.Category = "Drawing standards" OrElse i.Category = "Title block" Then Return True
        If i.ViewName = "<sheet>" Then Return Options.ScopeSheetNotes
        For Each name As String In i.ViewName.Split("|"c)
            For Each sh As String In i.Sheet.Split("|"c)
                If views.Contains(sh & "/" & name) Then Return True
            Next
        Next
        Return False
    End Function
    Private Sub SetLocation(i As Item, point As Point3d)
        i.X = point.X : i.Y = point.Y : i.Z = point.Z : i.HasLocation = True
    End Sub
    Private Sub Membership(i As Item, annotation As Annotation)
        Dim sheets As New SortedSet(Of String)(StringComparer.Ordinal), views As New SortedSet(Of String)(StringComparer.Ordinal)
        ' GetViews is valid for PMI display instances, not general drafting ownership.
        ' First resolve native view dependence in this owner part's view map.
        Dim dependent As Integer = 0, nativeViewName As String = ""
        DrawingChecker.UF.View.AskViewDependentStatus(annotation.Tag, dependent, nativeViewName)
        If dependent = 1 Then
            Dim matched As New List(Of Tag)()
            For Each entry As KeyValuePair(Of Tag, String) In NativeViewNames
                If entry.Value = nativeViewName Then matched.Add(entry.Key)
            Next
            If matched.Count = 1 Then
                sheets.Add(ViewSheets(matched(0))) : views.Add(ViewNames(matched(0)))
            End If
        Else
            For Each view As NXOpen.View In annotation.GetViews()
                If ViewSheets.ContainsKey(view.Tag) Then
                    sheets.Add(ViewSheets(view.Tag)) : views.Add(ViewNames(view.Tag))
                End If
            Next
        End If
        i.Sheet = String.Join("|", New List(Of String)(sheets).ToArray())
        i.ViewName = String.Join("|", New List(Of String)(views).ToArray())
        If i.Sheet = "" Then Problem("UNVERIFIED", "Sheet/view ownership", i, "Native view dependence / PMI GetViews did not resolve to a unique drawing sheet/view. Possible PMI, inherited annotation or unsupported membership; kept in unresolved inventory.")
    End Sub
    Private Sub ExtractAnnotation(a As Annotation)
        Dim category As String = "Symbols", supported As Boolean = False
        If TypeOf a Is Dimension Then
            category = "Dimensions" : supported = Options.Categories.Contains("Dimensions") OrElse Options.Categories.Contains("Tolerances")
        ElseIf TypeOf a Is BaseDatum Then
            category = "Datums" : supported = Options.Categories.Contains("Datums") OrElse Options.Categories.Contains("GD&T")
        ElseIf TypeOf a Is BaseFcf Then
            category = "GD&T" : supported = Options.Categories.Contains("GD&T")
        ElseIf TypeOf a Is Note OrElse TypeOf a Is NXOpen.Annotations.Label Then
            category = "Notes" : supported = Options.Categories.Contains("Notes") OrElse Options.Categories.Contains("Spelling")
        End If
        If category <> "Symbols" AndAlso Not supported Then Return
        Dim i As Item = Nothing
        Try
            i = BaseItem(a, category)
            Membership(i, a)
            If Not InScope(i) Then Return
            Data.Items.Add(i)
            SetLocation(i, a.AnnotationOrigin)
            i.Put("Suppressed", a.Suppressed) : i.Put("Blanked", a.IsBlanked)
            i.Put("Retained", a.IsRetained) : i.Put("OutOfDate", a.IsOutOfDate)
            If a.IsRetained OrElse a.IsOutOfDate Then Problem("WARNING", "Annotation state", i, "Retained or out-of-date annotation; geometry/reference validity needs review. No update is performed.")
            If TypeOf a Is Dimension Then
                ExtractDimension(CType(a, Dimension), i)
            ElseIf TypeOf a Is BaseDatum Then
                i.Put("Datum identifier", CType(a, BaseDatum).Label)
                i.RawText = CType(a, BaseDatum).Label
                ExtractAttachments(a, i)
            ElseIf TypeOf a Is BaseFcf Then
                ExtractFcf(CType(a, BaseFcf), i)
                ExtractAttachments(a, i)
            ElseIf TypeOf a Is SimpleDraftingAid Then
                i.RawText = JoinLines(CType(a, SimpleDraftingAid).GetText())
                If category = "Symbols" Then Problem("UNVERIFIED", "Symbol semantics", i, "Raw drafting text is captured; symbol structure and engineering meaning are not decoded.")
            Else
                Problem("UNVERIFIED", "Unsupported annotation", i, "Annotation type " & i.Kind & " is inventoried but has no semantic extractor.")
            End If
        Catch ex As Exception
            Problem("ERROR", "Annotation extraction " & a.Tag.ToString(), i, ex.GetType().Name & ": " & ex.Message)
        End Try
    End Sub
    Private Sub ExtractDimension(dimension As Dimension, i As Item)
        Dim main() As String = Nothing, dual() As String = Nothing
        dimension.GetDimensionText(main, dual)
        i.RawText = JoinLines(main) & If(JoinLines(dual) = "", "", vbLf & "DUAL: " & JoinLines(dual))
        i.Put("Measurement type", dimension.GetMeasurementType().ToString())
        i.NumericUnit = "NX dimension API / " & dimension.GetMeasurementType().ToString()
        i.Put("Value-unit validation", "Raw computed size and tolerance API values; confirm length/angular/display units with NX fixture before engineering interpretation.")
        If Options.Categories.Contains("Dimensions") Then
            i.Numbers("Computed nominal") = dimension.ComputedSize
            i.Put("Reference flag", dimension.ReferenceDimensionFlag)
            i.Put("Inspection flag", dimension.InspectionDimensionFlag)
            i.Put("Format.Nominal precision", dimension.NominalDecimalPlaces)
            Using appended As AppendedText = dimension.GetAppendedText()
                i.Put("Appended.Before", JoinLines(appended.GetBeforeText()))
                i.Put("Appended.After", JoinLines(appended.GetAfterText()))
                i.Put("Appended.Above", JoinLines(appended.GetAboveText()))
                i.Put("Appended.Below", JoinLines(appended.GetBelowText()))
            End Using
        Else
            i.RawText = ""
        End If
        If Options.Categories.Contains("Tolerances") Then
            i.Put("Tolerance type", dimension.ToleranceType.ToString())
            i.Numbers("Upper API tolerance") = dimension.UpperToleranceValue
            i.Numbers("Lower API tolerance") = dimension.LowerToleranceValue
            i.Put("Format.Tolerance precision", dimension.ToleranceDecimalPlaces)
            i.Put("Fit.Hole deviation", dimension.LimitFitDeviation)
            i.Put("Fit.Hole grade", dimension.LimitFitGrade)
            i.Put("Fit.Shaft deviation", dimension.LimitFitShaftDeviation)
            i.Put("Fit.Shaft grade", dimension.LimitFitShaftGrade)
            i.Put("Fit.Display style", dimension.LimitFitDisplayStyle.ToString())
        End If
        Using preferences As DimensionPreferences = dimension.GetDimensionPreferences()
            Using units As UnitsFormatPreferences = preferences.GetUnitsFormatPreferences()
                i.Put("Primary display unit", units.PrimaryDimensionUnit.ToString())
                i.Put("Dual display unit", units.DualDimensionUnit.ToString())
                i.Put("Format.Angular format", units.AngularDimensionFormat.ToString())
            End Using
        End Using
        ExtractAttachments(dimension, i)
        If i.Kind.IndexOf("Chamfer", StringComparison.OrdinalIgnoreCase) >= 0 OrElse i.Kind.IndexOf("Hole", StringComparison.OrdinalIgnoreCase) >= 0 OrElse i.Kind.IndexOf("Ordinate", StringComparison.OrdinalIgnoreCase) >= 0 Then Problem("UNVERIFIED", "Compound dimension semantics", i, "Base dimension data and displayed text captured. Compound values, hole/thread parameters, origin, baseline/chain grouping need subtype-specific NX 2306 validation.")
    End Sub
    Private Sub ExtractAttachments(a As Annotation, i As Item)
        ' Geometric references are diagnostic evidence, never universal cross-part identity.
        i.Put("Associativity count", a.NumberOfAssociativities)
        Dim refs As New List(Of String)()
        If TypeOf a Is Dimension Then
            Dim dimension As Dimension = CType(a, Dimension)
            If a.NumberOfAssociativities > 0 Then AppendEvidence(dimension.GetFirstAssociativity(), refs)
            If a.NumberOfAssociativities > 1 Then AppendEvidence(dimension.GetSecondAssociativity(), refs)
            If a.NumberOfAssociativities > 2 Then Problem("UNVERIFIED", "Additional associativities", i, "Only documented first/second dimension references are extracted in this stage.")
        ElseIf a.NumberOfAssociativities > 0 Then
            Problem("UNVERIFIED", "Attachment extraction", i, "Associativity count captured; indexed datum/FCF attachment interpretation needs a recorded NX 2306 fixture. No attachment equality is assumed.")
        End If
        ' Not included in ordinary field equality: journal refs are scoped to each part.
        i.Put("Format.Attachment journal evidence", String.Join("|", refs.ToArray()))
    End Sub
    Private Sub AppendEvidence(assoc As Associativity, refs As List(Of String))
        If assoc Is Nothing Then Return
        Using assoc
            If assoc.FirstObject IsNot Nothing Then refs.Add(assoc.FirstObject.GetType().Name & ":" & assoc.FirstObject.JournalIdentifier)
            If assoc.SecondObject IsNot Nothing Then refs.Add(assoc.SecondObject.GetType().Name & ":" & assoc.SecondObject.JournalIdentifier)
        End Using
    End Sub
    Private Sub ExtractFcf(fcf As BaseFcf, i As Item)
        Dim builder As FeatureControlFrameBuilder = Nothing
        Try
            If TypeOf fcf Is Fcf Then
                builder = Owner.Annotations.CreatePmiFeatureControlFrameBuilder(CType(fcf, Fcf))
            Else
                builder = Owner.Annotations.CreateDraftingFeatureControlFrameBuilder(fcf)
            End If
            i.Put("Characteristic", builder.Characteristic.ToString())
            i.Put("Frame style", builder.FrameStyle.ToString())
            Dim frames() As FeatureControlFrameDataBuilder = Nothing
            builder.FeatureControlFrameData(frames)
            If frames Is Nothing OrElse frames.Length = 0 Then Throw New Exception("FCF builder returned no frame data.")
            i.Put("Frame count", frames.Length)
            For index As Integer = 0 To frames.Length - 1
                Dim frame As FeatureControlFrameDataBuilder = frames(index)
                Dim prefix As String = "Frame[" & index.ToString() & "]."
                ' Whitelist of documented scalar properties. Reflection keeps optional properties
                ' from becoming compile-time dependencies; unavailable properties are errors.
                ReadScalars(frame, i, prefix, "ToleranceValue|MaterialModifier|ZoneShape|Projected|ProjectedValue|CommonZone|SeparateZone|FreeState|TangentPlane|StatisticalTolerance|MaximumTolerance|MaximumToleranceValue|UnitBasis|UnitBasisValue1|UnitBasisValue2|PrimaryDatumExtendedText|SecondaryDatumExtendedText|TertiaryDatumExtendedText|CircleU|CircleUvalue|DerivedFeature|DynamicProfile|LeastSquaresAssoc|LeastSquaresAssocExtConstraint|LeastSquaresAssocIntConstraint|LeastSquaresFeature|MaxInscribedAssoc|MaxInscribedFeature|MinCircumscribedAssoc|MinCircumscribedFeature|MinimaxAssoc|MinimaxAssocExtConstraint|MinimaxAssocIntConstraint|MinimaxFeature|OffsetZone|OrientationConstraint|PeakHeight|ReciprocityRequirement|StandardDeviation|TotalRangeOfDeviations|Uz|ValleyDepth|VariableAngle|AreaSymbol")
                ReadDatum(frame.PrimaryDatumReference, i, prefix & "Primary.")
                ReadDatum(frame.SecondaryDatumReference, i, prefix & "Secondary.")
                ReadDatum(frame.TertiaryDatumReference, i, prefix & "Tertiary.")
                ReadCompound(frame.PrimaryCompoundDatumReference, i, prefix & "PrimaryCommon.")
                ReadCompound(frame.SecondaryCompoundDatumReference, i, prefix & "SecondaryCommon.")
                ReadCompound(frame.TertiaryCompoundDatumReference, i, prefix & "TertiaryCommon.")
            Next
            i.RawText = JoinLines(fcf.GetText())
            Problem("UNVERIFIED", "FCF extended semantics", i, "Structured characteristic, ordered frames, datums, zone/modifier fields captured. Indicator lists, business modifiers, extended symbol text and complete ASME/ISO semantic validity are not certified by this version.")
        Finally
            If builder IsNot Nothing Then builder.Destroy() ' Never Commit.
        End Try
    End Sub
    Private Sub ReadScalars(obj As Object, i As Item, prefix As String, whitelist As String)
        For Each propertyName As String In whitelist.Split("|"c)
            Try
                Dim info As PropertyInfo = obj.GetType().GetProperty(propertyName)
                If info Is Nothing OrElse Not info.CanRead Then Throw New MissingMemberException(obj.GetType().Name, propertyName)
                i.Put(prefix & propertyName, info.GetValue(obj, Nothing))
            Catch ex As Exception
                Problem("ERROR", "Read " & prefix & propertyName, i, ex.Message)
            End Try
        Next
    End Sub
    Private Sub ReadDatum(datum As DatumReferenceBuilder, i As Item, prefix As String)
        If datum Is Nothing Then
            i.Put(prefix & "Letter", "") : Return
        End If
        ReadScalars(datum, i, prefix, "Letter|MaterialCondition|FreeState|Projected")
        If datum.Letter <> "" Then i.Datums.Add(datum.Letter)
    End Sub
    Private Sub ReadCompound(compound As CompoundDatumReferenceBuilder, i As Item, prefix As String)
        If compound Is Nothing Then
            i.Put(prefix & "Count", 0) : Return
        End If
        Dim entries() As DatumReferenceBuilder = compound.DatumReferenceList.GetContents()
        i.Put(prefix & "Count", entries.Length)
        For index As Integer = 0 To entries.Length - 1
            ReadDatum(entries(index), i, prefix & "[" & index.ToString() & "].")
        Next
        If entries.Length > 0 Then Problem("UNVERIFIED", "Common datum interpretation", i, "Ordered common-datum members captured; compound extended text and standard-specific precedence require engineering review.")
    End Sub
    Private Sub ExtractView(view As DraftingView, sheet As DrawingSheet)
        Dim i As Item = Nothing
        Try
            i = BaseItem(view, "Views") : i.Sheet = sheet.Name : i.ViewName = view.Name
            If Not InScope(i) Then Return
            Data.Items.Add(i) : SetLocation(i, view.GetDrawingReferencePoint())
            i.Numbers("Scale") = view.Style.General.Scale
            i.Numbers("Angle") = view.Style.General.Angle
            i.Put("Broken", view.IsBroken) : i.Put("OutOfDate", view.IsOutOfDate)
            i.Put("Hidden lines", view.Style.HiddenLines.Hiddenline)
            i.Put("View label enabled", view.Style.General.ViewLabel)
            i.Put("Scale label enabled", view.Style.General.ScaleLabel)
            Dim m As Matrix3x3 = view.Matrix
            i.Put("Orientation matrix", String.Join(";", New String() {Num(m.Xx), Num(m.Xy), Num(m.Xz), Num(m.Yx), Num(m.Yy), Num(m.Yz), Num(m.Zx), Num(m.Zy), Num(m.Zz)}))
            If view.IsOutOfDate Then Problem("WARNING", "View out of date", i, "View is not updated; generated geometry may not represent the model.")
        Catch ex As Exception
            Problem("ERROR", "View extraction", i, ex.Message)
        End Try
    End Sub
    Private Sub ExtractSheet(sheet As DrawingSheet)
        Dim i As Item = BaseItem(sheet, "Drawing standards")
        i.Sheet = sheet.Name : i.Kind = "DrawingSheet"
        If Not InScope(i) Then Return
        Data.Items.Add(i)
        i.Numbers("Length") = sheet.Length : i.Numbers("Height") = sheet.Height
        i.Put("Sheet units", sheet.Units.ToString()) : i.Put("Projection", sheet.ProjectionAngle.ToString())
        Dim numerator As Double = 0, denominator As Double = 0
        sheet.GetScale(numerator, denominator)
        i.Put("Scale numerator", Num(numerator)) : i.Put("Scale denominator", Num(denominator))
        i.Put("Sheet count in part", Owner.DrawingSheets.ToArray().Length)
    End Sub
    Private Sub ExtractAttributes()
        Dim profile As Dictionary(Of String, String) = Profiles.Read(Options.ProfilePath)
        Dim count As Integer = 0
        For Each entry As KeyValuePair(Of String, String) In profile
            If Not entry.Key.StartsWith("Attribute.", StringComparison.Ordinal) Then Continue For
            Dim i As Item = BaseItem(Owner, "Title block")
            i.Kind = "MappedPartAttribute" : i.Key = entry.Key : i.CheckId = entry.Key
            i.Name = entry.Key.Substring(10) : i.JournalId = entry.Key
            Data.Items.Add(i) : count += 1
            If Owner.HasUserAttribute(entry.Value, NXObject.AttributeType.String, -1) Then
                i.RawText = Owner.GetStringUserAttribute(entry.Value, -1)
                i.Put("Attribute title", entry.Value)
                If i.RawText.Trim() = "" Then Problem("WARNING", "Blank mapped attribute", i, entry.Value & " is blank.")
            Else
                Problem("UNVERIFIED", "Missing mapped string attribute", i, entry.Value & " is unavailable; no title-block cell value is inferred.")
            End If
        Next
        If count = 0 Then Problem("UNVERIFIED", "Title-block mapping", Nothing, "No Attribute.* mappings configured. Title-block tables, template links and Teamcenter values need a site-specific adapter.")
    End Sub
    Private Sub ValidateDatums()
        If Not Options.Categories.Contains("Datums") AndAlso Not Options.Categories.Contains("GD&T") Then Return
        Dim definitions As New Dictionary(Of String, List(Of Item))(StringComparer.Ordinal)
        For Each i As Item In Data.Items
            If i.Category <> "Datums" OrElse Not i.Fields.ContainsKey("Datum identifier") Then Continue For
            Dim name As String = i.Fields("Datum identifier")
            If Not definitions.ContainsKey(name) Then definitions(name) = New List(Of Item)()
            definitions(name).Add(i)
        Next
        For Each entry As KeyValuePair(Of String, List(Of Item)) In definitions
            If entry.Key.Trim() = "" Then Problem("WARNING", "Blank datum identifier", entry.Value(0), "Blank datum label.")
            If entry.Value.Count > 1 Then Problem("WARNING", "Repeated datum identifier", entry.Value(0), entry.Key & " occurs " & entry.Value.Count.ToString() & " times. Repetition across views/sheets can be legitimate; inspect attachment and definition rules.")
        Next
        For Each i As Item In Data.Items
            For Each datum As String In i.Datums
                If Not Regex.IsMatch(datum, "^[A-Z][A-Z0-9]*$") Then
                    Problem("UNVERIFIED", "Datum expression", i, "No plain-identifier interpretation for '" & datum & "'.")
                ElseIf Not definitions.ContainsKey(datum) Then
                    Problem("UNVERIFIED", "Datum definition unresolved", i, "Datum '" & datum & "' has no extracted definition in selected scope. It could be outside scope, PMI or unsupported; absence is not assumed.")
                End If
            Next
        Next
    End Sub
    Private Sub DuplicateIds()
        Dim ids As New Dictionary(Of String, Integer)(StringComparer.Ordinal)
        For Each i As Item In Data.Items
            If i.CheckId = "" Then Continue For
            Dim key As String = i.Category & ":" & i.CheckId
            If Not ids.ContainsKey(key) Then ids(key) = 0
            ids(key) += 1
        Next
        For Each i As Item In Data.Items
            If i.CheckId <> "" AndAlso ids(i.Category & ":" & i.CheckId) > 1 Then Problem("UNVERIFIED", "Duplicate CHECK_ID", i, "CHECK_ID must be unique within its category; identity cannot be trusted.")
        Next
    End Sub
    Private Sub AddCoverage()
        Problem("UNVERIFIED", "Release coverage", Nothing, "Staged implementation: drawing geometry, complete GD&T semantics, attachment equivalence, displayed overrides, table cells, welding/surface-finish structures, section arrows, view alignment, borders/zones and revision-history semantics need validated adapters. No engineering-release approval is issued.")
        Problem("UNVERIFIED", "Sheet membership coverage", Nothing, "Drafting ownership uses UF.View.AskViewDependentStatus with owner-part view names; GetViews is only a PMI fallback. Unresolved annotations stay in inventory; hidden/inherited/PMI/custom entities may require UF drafting ownership adapters. Scope completeness needs the local validation matrix.")
        If Options.Categories.Contains("Drawing standards") Then Problem("UNVERIFIED", "Standards certification", Nothing, "Only configured company-profile rules and extracted metadata are evaluated. No unconfigured ASME/ISO rules are assumed.")
    End Sub
    Public Shared Function JoinLines(lines() As String) As String
        Return If(lines Is Nothing, "", String.Join(vbLf, lines))
    End Function
    Private Shared Function Num(value As Double) As String
        Return value.ToString("G17", CultureInfo.InvariantCulture)
    End Function
End Class

' -------------------- Explicit company rules / optional vocabulary --------------------
Public Class Profiles
    Public Shared Function Read(filename As String) As Dictionary(Of String, String)
        Dim profile As New Dictionary(Of String, String)(StringComparer.Ordinal)
        If filename.Trim() = "" Then Return profile
        For Each raw As String In File.ReadAllLines(filename, Encoding.UTF8)
            Dim line As String = raw.Trim()
            If line = "" OrElse line.StartsWith("#", StringComparison.Ordinal) Then Continue For
            Dim equals As Integer = line.IndexOf("="c)
            If equals < 1 Then Throw New FormatException("Profile lines must be Key=Value: " & line)
            Dim key As String = line.Substring(0, equals).Trim(), value As String = line.Substring(equals + 1).Trim()
            If profile.ContainsKey(key) Then Throw New FormatException("Duplicate profile key: " & key)
            If Not key.StartsWith("Attribute.", StringComparison.Ordinal) AndAlso Not key.StartsWith("Expected.", StringComparison.Ordinal) AndAlso key <> "AllowedSheetSizesMm" AndAlso key <> "ExpectedProjection" Then Throw New FormatException("Unsupported profile key: " & key)
            profile.Add(key, value)
        Next
        Return profile
    End Function
    Public Shared Sub Apply(current As Snapshot, options As Settings, rows As List(Of Finding))
        Dim config As Dictionary(Of String, String) = Read(options.ProfilePath)
        Dim applied As Integer = 0
        If options.Categories.Contains("Drawing standards") OrElse options.Categories.Contains("Title block") Then
            For Each i As Item In current.Items
                If i.Category = "Title block" AndAlso config.ContainsKey("Expected." & i.Name) Then
                    applied += 1
                    rows.Add(Inspector.Make(If(i.State <> "OK", "UNVERIFIED", If(i.RawText = config("Expected." & i.Name), "PASS", "FAIL")), "Title block", "Configured value " & i.Name, Nothing, i, config("Expected." & i.Name), i.RawText, "Explicit user company-profile rule; not a standards certification."))
                End If
                If i.Category = "Drawing standards" AndAlso config.ContainsKey("ExpectedProjection") Then
                    applied += 1
                    Dim value As String = If(i.Fields.ContainsKey("Projection"), i.Fields("Projection"), "<unavailable>")
                    rows.Add(Inspector.Make(If(value = config("ExpectedProjection"), "PASS", "FAIL"), "Drawing standards", "Configured projection", Nothing, i, config("ExpectedProjection"), value, "Exact NX projection enum name from company profile."))
                End If
                If i.Category = "Drawing standards" AndAlso config.ContainsKey("AllowedSheetSizesMm") Then
                    applied += 1
                    Dim w As Double = i.Numbers("Length"), h As Double = i.Numbers("Height")
                    If i.Fields("Sheet units").IndexOf("Inch", StringComparison.OrdinalIgnoreCase) >= 0 Then
                        w *= 25.4 : h *= 25.4
                    End If
                    Dim match As Boolean = False
                    For Each token As String In config("AllowedSheetSizesMm").Split(";"c)
                        Dim values() As String = token.Trim().ToLowerInvariant().Split("x"c)
                        If values.Length <> 2 Then Throw New FormatException("AllowedSheetSizesMm must contain widthxheight pairs separated by semicolons.")
                        Dim a As Double = Double.Parse(values(0), CultureInfo.InvariantCulture), b As Double = Double.Parse(values(1), CultureInfo.InvariantCulture)
                        If a <= 0 OrElse b <= 0 Then Throw New FormatException("Sheet dimensions must be positive.")
                        If (Math.Abs(w - a) < 0.1 AndAlso Math.Abs(h - b) < 0.1) OrElse (Math.Abs(w - b) < 0.1 AndAlso Math.Abs(h - a) < 0.1) Then match = True
                    Next
                    rows.Add(Inspector.Make(If(match, "PASS", "FAIL"), "Drawing standards", "Configured sheet sizes", Nothing, i, config("AllowedSheetSizesMm"), w.ToString("G8", CultureInfo.InvariantCulture) & "x" & h.ToString("G8", CultureInfo.InvariantCulture) & " mm", "Explicit allowed-size list; both orientations accepted, 0.1 mm metadata comparison epsilon."))
                End If
            Next
            If applied = 0 Then rows.Add(Inspector.Make("UNVERIFIED", "Drawing standards", "No executable company rules", Nothing, Nothing, "", "", "Configure Expected.*, ExpectedProjection or AllowedSheetSizesMm. Selecting a mode does not create ASME/ISO rules."))
        End If
        If options.Categories.Contains("Spelling") Then Spell(current, options.DictionaryPath, rows)
    End Sub
    Private Shared Sub Spell(current As Snapshot, dictionaryPath As String, rows As List(Of Finding))
        If dictionaryPath.Trim() = "" Then
            rows.Add(Inspector.Make("UNVERIFIED", "Spelling", "Vocabulary missing", Nothing, Nothing, "", "", "Select an approved plain UTF-8 vocabulary file, one allowed word/token per line. No online spelling service is used.")) : Return
        End If
        Dim vocabulary As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
        For Each line As String In File.ReadAllLines(dictionaryPath, Encoding.UTF8)
            Dim word As String = line.Trim()
            If word <> "" AndAlso Not word.StartsWith("#", StringComparison.Ordinal) Then vocabulary.Add(word)
        Next
        If vocabulary.Count = 0 Then Throw New FormatException("The selected vocabulary file contains no words.")
        For Each i As Item In current.Items
            If i.Category <> "Notes" Then Continue For
            Dim unknown As New SortedSet(Of String)(StringComparer.OrdinalIgnoreCase)
            For Each m As Match In Regex.Matches(i.RawText, "[A-Za-z][A-Za-z'-]{2,}")
                If Not vocabulary.Contains(m.Value) Then unknown.Add(m.Value)
            Next
            If unknown.Count > 0 Then rows.Add(Inspector.Make("WARNING", "Spelling", "Vocabulary review", Nothing, i, "Approved vocabulary", String.Join(", ", New List(Of String)(unknown).ToArray()), "Unknown tokens may be names, abbreviations, NX text codes or valid technical terms; no automatic correction."))
        Next
    End Sub
End Class

' -------------------- WinForms setup, independent scope, result review --------------------
Public Class CheckerForm
    Inherits F.Form
    Private OwnerPart As Part, ReferencePart As Part
    Private Mode As New F.ComboBox(), Categories As New F.CheckedListBox()
    Private ReferencePath As New F.TextBox(), OutputPath As New F.TextBox(), ProfilePath As New F.TextBox(), WordsPath As New F.TextBox()
    Private ReferenceImage As New F.TextBox()
    Private Epsilon As New F.NumericUpDown(), Movement As New F.NumericUpDown()
    Private TrustIds As New F.CheckBox(), SheetNotes As New F.CheckBox(), Capture As New F.CheckBox()
    Private CS As New F.CheckedListBox(), RS As New F.CheckedListBox(), CV As New F.CheckedListBox(), RV As New F.CheckedListBox()
    Private AllCS As New F.CheckBox(), AllRS As New F.CheckBox(), AllCV As New F.CheckBox(), AllRV As New F.CheckBox()
    Private Status As New F.Label(), Grid As New F.DataGridView(), Filter As New F.ComboBox()
    Private ExtractButton As New F.Button(), CompareButton As New F.Button(), LoadButton As New F.Button(), ExportButton As New F.Button(), CancelButtonControl As New F.Button()
    Private Options As Settings, CurData As Snapshot, RefData As Snapshot
    Private Rows As New List(Of Finding)(), RunFolder As String = "", CurrentImage As String = ""
    Private ExtractionOnly As Boolean, Busy As Boolean, CancelRequested As Boolean
    Private SetupPanel As F.Panel, ReferenceImageAtRun As String = ""
    Public Sub New(p As Part)
        OwnerPart = p
        Text = "NX 2306 drawing checker - staged native inspection"
        Width = 1250 : Height = 880 : MinimumSize = New D.Size(1050, 740)
        StartPosition = F.FormStartPosition.CenterScreen
        Dim tabs As New F.TabControl With {.Dock = F.DockStyle.Fill}
        Dim setup As New F.TabPage("Setup and scope"), results As New F.TabPage("Results")
        tabs.TabPages.Add(setup) : tabs.TabPages.Add(results)
        Dim layout As New F.TableLayoutPanel With {.Dock = F.DockStyle.Top, .AutoSize = True, .ColumnCount = 3, .Padding = New F.Padding(12)}
        layout.ColumnStyles.Add(New F.ColumnStyle(F.SizeType.Absolute, 180))
        layout.ColumnStyles.Add(New F.ColumnStyle(F.SizeType.Percent, 100))
        layout.ColumnStyles.Add(New F.ColumnStyle(F.SizeType.Absolute, 120))
        Dim currentLabel As New F.Label With {.Text = p.FullPath, .AutoSize = True}
        AddRow(layout, "Current displayed/work part", currentLabel, Nothing)
        LoadButton.Text = "Load .prt"
        AddHandler LoadButton.Click, AddressOf LoadReference
        AddRow(layout, "Reference native drawing", ReferencePath, LoadButton)
        ReferencePath.ReadOnly = True
        AddRow(layout, "Output folder", OutputPath, BrowseButton(OutputPath, "folder"))
        OutputPath.Text = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments)
        AddRow(layout, "Company profile (optional)", ProfilePath, BrowseButton(ProfilePath, "profile"))
        AddRow(layout, "Vocabulary (optional)", WordsPath, BrowseButton(WordsPath, "words"))
        AddRow(layout, "Reference image (optional)", ReferenceImage, BrowseButton(ReferenceImage, "image"))
        Mode.DropDownStyle = F.ComboBoxStyle.DropDownList
        Mode.Items.AddRange(New Object() {"Engineering", "Exact", "Configured company rules"}) : Mode.SelectedIndex = 0
        AddRow(layout, "Comparison mode", Mode, Nothing)
        Epsilon.DecimalPlaces = 9 : Epsilon.Minimum = 0 : Epsilon.Maximum = 0.001D : Epsilon.Increment = 0.000001D : Epsilon.Value = 0.000001D
        AddRow(layout, "Numeric equality epsilon", Epsilon, New F.Label With {.Text = "native units", .AutoSize = True})
        Movement.DecimalPlaces = 2 : Movement.Minimum = 0 : Movement.Maximum = 1000 : Movement.Value = 1
        AddRow(layout, "Layout warning threshold", Movement, New F.Label With {.Text = "mm", .AutoSize = True})
        TrustIds.Text = "Trust journal identifiers only for known related revisions" : TrustIds.AutoSize = True
        SheetNotes.Text = "Include sheet-level notes when filtering drafting views" : SheetNotes.Checked = True : SheetNotes.AutoSize = True
        Capture.Text = "Capture current displayed graphics image on run (not all sheets)" : Capture.Checked = True : Capture.AutoSize = True
        AddRow(layout, "Matching / evidence", TrustIds, Nothing) : AddRow(layout, "", SheetNotes, Nothing) : AddRow(layout, "", Capture, Nothing)
        Categories.CheckOnClick = True : Categories.Height = 92 : Categories.MultiColumn = True : Categories.ColumnWidth = 155
        For Each category As String In New String() {"GD&T", "Datums", "Dimensions", "Tolerances", "Notes", "Views", "Title block", "Drawing standards", "Spelling"}
            Categories.Items.Add(category, category <> "Spelling")
        Next
        AddRow(layout, "Inspection categories", Categories, Nothing)
        Dim scopes As New F.TableLayoutPanel With {.Dock = F.DockStyle.Top, .Height = 220, .ColumnCount = 4}
        For n As Integer = 0 To 3
            scopes.ColumnStyles.Add(New F.ColumnStyle(F.SizeType.Percent, 25))
        Next
        scopes.Controls.Add(ScopePanel("Current sheets", AllCS, CS), 0, 0)
        scopes.Controls.Add(ScopePanel("Reference sheets", AllRS, RS), 1, 0)
        scopes.Controls.Add(ScopePanel("Current views", AllCV, CV), 2, 0)
        scopes.Controls.Add(ScopePanel("Reference views", AllRV, RV), 3, 0)
        PopulateScope(p, CS, CV)
        Dim note As New F.Label With {.Dock = F.DockStyle.Top, .Height = 54, .Text = "Native NX only. PDF/DXF/CGM parsing, complete standards certification and geometry-based feature equivalence are not implemented. Unknown ownership and unsupported objects are reported. Source parts are never saved/updated/closed. Reference remains open."}
        Dim container As New F.Panel With {.Dock = F.DockStyle.Fill, .AutoScroll = True}
        SetupPanel = container
        container.Controls.Add(note) : container.Controls.Add(scopes) : container.Controls.Add(layout)
        setup.Controls.Add(container)
        Grid.Dock = F.DockStyle.Fill : Grid.ReadOnly = True : Grid.AllowUserToAddRows = False
        Grid.SelectionMode = F.DataGridViewSelectionMode.FullRowSelect : Grid.MultiSelect = False
        Grid.AutoSizeColumnsMode = F.DataGridViewAutoSizeColumnsMode.Fill : Grid.RowHeadersVisible = False
        For Each name As String In New String() {"ID", "Status", "Category", "Check", "Reference", "Current", "Message"}
            Grid.Columns.Add(name, name)
        Next
        Grid.Columns("ID").FillWeight = 50 : Grid.Columns("Status").FillWeight = 40
        Grid.DefaultCellStyle.WrapMode = F.DataGridViewTriState.True : Grid.AutoSizeRowsMode = F.DataGridViewAutoSizeRowsMode.AllCells
        Dim resultTools As New F.FlowLayoutPanel With {.Dock = F.DockStyle.Top, .Height = 38}
        Filter.DropDownStyle = F.ComboBoxStyle.DropDownList
        Filter.Items.AddRange(New Object() {"All", "Non-PASS", "FAIL", "MISSING", "EXTRA", "WARNING", "UNVERIFIED", "ERROR", "PASS"}) : Filter.SelectedIndex = 0
        AddHandler Filter.SelectedIndexChanged, Sub(sender, args) RefreshGrid()
        resultTools.Controls.Add(Filter)
        Dim highlight As New F.Button With {.Text = "Highlight current", .AutoSize = True}
        AddHandler highlight.Click, AddressOf HighlightCurrent
        resultTools.Controls.Add(highlight)
        Dim clear As New F.Button With {.Text = "Clear highlight", .AutoSize = True}
        AddHandler clear.Click, Sub(sender, args) DrawingChecker.ClearHighlight()
        resultTools.Controls.Add(clear)
        resultTools.Controls.Add(New F.Label With {.Text = "Open the reported sheet manually; view/sheet settings are preserved.", .AutoSize = True, .Padding = New F.Padding(5, 8, 0, 0)})
        results.Controls.Add(Grid) : results.Controls.Add(resultTools)
        Dim actions As New F.FlowLayoutPanel With {.Dock = F.DockStyle.Bottom, .Height = 46}
        ExtractButton.Text = "Extract current" : CompareButton.Text = "Compare" : ExportButton.Text = "Export Excel" : CancelButtonControl.Text = "Cancel scan"
        ExtractButton.AutoSize = True : CompareButton.AutoSize = True : ExportButton.AutoSize = True : CancelButtonControl.AutoSize = True
        ExportButton.Enabled = False : CancelButtonControl.Enabled = False
        AddHandler ExtractButton.Click, Sub(sender, args) Run(True, tabs)
        AddHandler CompareButton.Click, Sub(sender, args) Run(False, tabs)
        AddHandler ExportButton.Click, AddressOf Export
        AddHandler CancelButtonControl.Click, Sub(sender, args) CancelRequested = True
        actions.Controls.AddRange(New F.Control() {ExtractButton, CompareButton, ExportButton, CancelButtonControl})
        Status.Dock = F.DockStyle.Bottom : Status.Height = 36 : Status.Text = "Ready. First validate on copied test drawings."
        Controls.Add(tabs) : Controls.Add(Status) : Controls.Add(actions)
        AddHandler FormClosing, AddressOf ClosingForm
    End Sub
    Private Sub AddRow(panel As F.TableLayoutPanel, caption As String, control As F.Control, extra As F.Control)
        Dim row As Integer = panel.RowCount
        panel.RowCount += 1 : panel.RowStyles.Add(New F.RowStyle(F.SizeType.AutoSize))
        panel.Controls.Add(New F.Label With {.Text = caption, .AutoSize = True, .Padding = New F.Padding(0, 6, 0, 0)}, 0, row)
        control.Dock = F.DockStyle.Fill : panel.Controls.Add(control, 1, row)
        If extra IsNot Nothing Then
            extra.Dock = F.DockStyle.Fill : panel.Controls.Add(extra, 2, row)
        End If
    End Sub
    Private Function BrowseButton(target As F.TextBox, kind As String) As F.Button
        Dim button As New F.Button With {.Text = "Browse...", .AutoSize = True}
        AddHandler button.Click, Sub(sender, args)
            If kind = "folder" Then
                Using dialog As New F.FolderBrowserDialog()
                    If dialog.ShowDialog() = F.DialogResult.OK Then target.Text = dialog.SelectedPath
                End Using
            Else
                Using dialog As New F.OpenFileDialog()
                    dialog.Filter = If(kind = "image", "Image|*.png;*.jpg;*.jpeg", "Text/configuration|*.ini;*.txt|All files|*.*")
                    If dialog.ShowDialog() = F.DialogResult.OK Then target.Text = dialog.FileName
                End Using
            End If
        End Sub
        Return button
    End Function
    Private Function ScopePanel(title As String, all As F.CheckBox, list As F.CheckedListBox) As F.Control
        Dim group As New F.GroupBox With {.Text = title, .Dock = F.DockStyle.Fill}
        all.Text = "All" : all.Checked = True : all.Dock = F.DockStyle.Top
        list.Dock = F.DockStyle.Fill : list.CheckOnClick = True : list.Enabled = False
        AddHandler all.CheckedChanged, Sub(sender, args) list.Enabled = Not all.Checked
        group.Controls.Add(list) : group.Controls.Add(all)
        Return group
    End Function
    Private Sub PopulateScope(p As Part, sheets As F.CheckedListBox, views As F.CheckedListBox)
        sheets.Items.Clear() : views.Items.Clear()
        For Each sheet As DrawingSheet In p.DrawingSheets.ToArray()
            sheets.Items.Add(sheet.Name, True)
            For Each view As DraftingView In sheet.GetDraftingViews()
                views.Items.Add(sheet.Name & "/" & view.Name, True)
            Next
        Next
    End Sub
    Private Sub LoadReference(sender As Object, args As EventArgs)
        If Busy Then Return
        Using dialog As New F.OpenFileDialog()
            dialog.Filter = "Native NX drawing part|*.prt"
            If dialog.ShowDialog() <> F.DialogResult.OK Then Return
            Try
                Dim p As Part = Nothing
                For Each loaded As BasePart In DrawingChecker.Session.Parts
                    If String.Equals(loaded.FullPath, dialog.FileName, StringComparison.OrdinalIgnoreCase) Then p = TryCast(loaded, Part)
                Next
                If p Is Nothing Then
                    Dim loadStatus As PartLoadStatus = Nothing
                    Try
                        p = DrawingChecker.Session.Parts.Open(dialog.FileName, loadStatus)
                        If loadStatus IsNot Nothing AndAlso loadStatus.NumberUnloadedParts > 0 Then DrawingChecker.Log("WARNING reference load: " & loadStatus.NumberUnloadedParts.ToString() & " unloaded parts; see extraction coverage.")
                    Finally
                        If loadStatus IsNot Nothing Then loadStatus.Dispose()
                    End Try
                End If
                If p Is Nothing OrElse p.DrawingSheets.ToArray().Length = 0 Then Throw New Exception("Selected part has no drawing sheets.")
                If p.Tag = OwnerPart.Tag Then Throw New Exception("Select a different reference drawing.")
                ReferencePart = p : ReferencePath.Text = p.FullPath : ReferencePath.ReadOnly = True
                PopulateScope(p, RS, RV)
                DrawingChecker.Log("Reference opened without display/work switch: " & p.FullPath)
                Status.Text = "Reference loaded. It will remain open after this journal exits."
            Catch ex As Exception
                F.MessageBox.Show(ex.Message, "Load reference", F.MessageBoxButtons.OK, F.MessageBoxIcon.Error)
            End Try
        End Using
    End Sub
    Private Function ReadOptions() As Settings
        Dim o As New Settings With {.Mode = Mode.Text, .NumericEpsilon = CDbl(Epsilon.Value), .LayoutMm = CDbl(Movement.Value), .TrustJournalIds = TrustIds.Checked, .ScopeSheetNotes = SheetNotes.Checked, .ProfilePath = ProfilePath.Text.Trim(), .DictionaryPath = WordsPath.Text.Trim(), .CurrentAllSheets = AllCS.Checked, .ReferenceAllSheets = AllRS.Checked, .CurrentAllViews = AllCV.Checked, .ReferenceAllViews = AllRV.Checked}
        For Each value As Object In Categories.CheckedItems
            o.Categories.Add(CStr(value))
        Next
        For Each value As Object In CS.CheckedItems
            o.CurrentSheets.Add(CStr(value))
        Next
        For Each value As Object In RS.CheckedItems
            o.ReferenceSheets.Add(CStr(value))
        Next
        For Each value As Object In CV.CheckedItems
            o.CurrentViews.Add(CStr(value))
        Next
        For Each value As Object In RV.CheckedItems
            o.ReferenceViews.Add(CStr(value))
        Next
        If o.Categories.Count = 0 Then Throw New Exception("Select at least one inspection category.")
        If Not o.CurrentAllSheets AndAlso o.CurrentSheets.Count = 0 Then Throw New Exception("Select at least one current sheet.")
        If Not o.ReferenceAllSheets AndAlso o.ReferenceSheets.Count = 0 Then Throw New Exception("Select at least one reference sheet.")
        If Not o.CurrentAllViews AndAlso o.CurrentViews.Count = 0 Then Throw New Exception("Select at least one current view.")
        If Not o.ReferenceAllViews AndAlso o.ReferenceViews.Count = 0 Then Throw New Exception("Select at least one reference view.")
        Profiles.Read(o.ProfilePath) ' Validate before scanning.
        Return o
    End Function
    Private Sub Progress(text As String)
        Status.Text = text : Status.Refresh()
        F.Application.DoEvents() ' UI only; NX access remains on this journal thread.
    End Sub
    Private Sub Run(extractOnly As Boolean, tabs As F.TabControl)
        If Busy Then Return
        Try
            If DrawingChecker.Session.Parts.Work.Tag <> OwnerPart.Tag OrElse DrawingChecker.Session.Parts.Display.Tag <> OwnerPart.Tag Then Throw New Exception("Restore the original displayed/work drawing before running.")
            Options = ReadOptions()
            If Not extractOnly AndAlso ReferencePart Is Nothing Then Throw New Exception("Load a native reference .prt first.")
            If Not Directory.Exists(OutputPath.Text) Then Throw New Exception("Select an existing output folder.")
            Busy = True : CancelRequested = False : SetBusy(True)
            DrawingChecker.ClearHighlight()
            RunFolder = Path.Combine(Path.GetFullPath(OutputPath.Text), "NX_Drawing_Check_" & DateTime.Now.ToString("yyyyMMdd_HHmmss") & "_" & Guid.NewGuid().ToString("N").Substring(0, 6))
            Directory.CreateDirectory(RunFolder)
            ExtractionOnly = extractOnly : CurrentImage = "" : ReferenceImageAtRun = ReferenceImage.Text.Trim() : ExportButton.Enabled = False
            DrawingChecker.Log("Mode=" & Options.Mode & "; epsilon=" & Options.NumericEpsilon.ToString("R", CultureInfo.InvariantCulture) & "; layout mm=" & Options.LayoutMm.ToString(CultureInfo.InvariantCulture))
            CurData = New Extractor(OwnerPart, Options, "Current", AddressOf Progress, Function() CancelRequested, False).Run()
            RefData = New Snapshot With {.Label = "Reference", .PartPath = "<not selected>", .PartUnit = ""}
            If Not extractOnly Then RefData = New Extractor(ReferencePart, Options, "Reference", AddressOf Progress, Function() CancelRequested, True).Run()
            If CancelRequested Then Throw New OperationCanceledException()
            If extractOnly Then
                Rows = New List(Of Finding)(CurData.Issues)
                For Each i As Item In CurData.Items
                    Rows.Add(Inspector.Make("UNVERIFIED", i.Category, "Extracted - not compared", Nothing, i, "", i.Describe(), "Extraction inventory only; this is not an inspection PASS."))
                Next
            Else
                Progress("Comparing structured records...")
                Rows = Inspector.Compare(RefData, CurData, Options, AddressOf Progress, Function() CancelRequested)
            End If
            Profiles.Apply(CurData, Options, Rows)
            If Capture.Checked Then
                Try
                    CurrentImage = Path.Combine(RunFolder, "Current_Active_Graphics.png")
                    CaptureCurrent(CurrentImage)
                Catch ex As Exception
                    CurrentImage = ""
                    Rows.Add(Inspector.Make("ERROR", "Images", "Current graphics capture", Nothing, Nothing, "", "", ex.Message))
                End Try
            End If
            Rows.Add(Inspector.Make("UNVERIFIED", "Images", "Image provenance / registration", Nothing, Nothing, ReferenceImageAtRun, CurrentImage, "Current PNG is the displayed graphics area only. Optional reference image is supplied by user. Neither is automatically registered to all sheets; no pixel difference markers are claimed."))
            Reports.WriteCsv(RunFolder, RefData, CurData, Rows, Options, ExtractionOnly)
            File.WriteAllText(Path.Combine(RunFolder, "Run_Log.txt"), String.Join(vbCrLf, DrawingChecker.Diagnostics.ToArray()), New UTF8Encoding(False))
            Status.Text = Inspector.Overall(Rows, ExtractionOnly) & " | " & Rows.Count.ToString() & " checks | CSV: " & RunFolder
            DrawingChecker.Log(Status.Text)
            RefreshGrid() : tabs.SelectedIndex = 1 : ExportButton.Enabled = True
        Catch ex As OperationCanceledException
            Rows.Clear() : ExportButton.Enabled = False : RefreshGrid()
            Status.Text = "Cancelled. No completed inspection report; source drawings unchanged."
            DrawingChecker.Log(Status.Text)
        Catch ex As Exception
            ExportButton.Enabled = False
            DrawingChecker.Log("ERROR run: " & ex.ToString())
            Status.Text = "ERROR - no completed inspection: " & ex.Message
            F.MessageBox.Show(ex.Message, "Drawing checker", F.MessageBoxButtons.OK, F.MessageBoxIcon.Error)
        Finally
            Busy = False : SetBusy(False)
        End Try
    End Sub
    Private Sub SetBusy(value As Boolean)
        ExtractButton.Enabled = Not value : CompareButton.Enabled = Not value : LoadButton.Enabled = Not value
        CancelButtonControl.Enabled = value
        SetupPanel.Enabled = Not value
    End Sub
    Private Sub RefreshGrid()
        Grid.Rows.Clear()
        For Each row As Finding In Rows
            If Filter.Text <> "All" AndAlso Not (Filter.Text = "Non-PASS" AndAlso row.Status <> "PASS") AndAlso Filter.Text <> row.Status Then Continue For
            Dim n As Integer = Grid.Rows.Add(row.ID, row.Status, row.Category, row.Rule, row.Reference, row.CurrentValue, row.Message)
            Grid.Rows(n).Tag = row
            Grid.Rows(n).DefaultCellStyle.BackColor = Reports.StatusColour(row.Status)
        Next
    End Sub
    Private Sub HighlightCurrent(sender As Object, args As EventArgs)
        If Busy OrElse Grid.CurrentRow Is Nothing Then Return
        Try
            DrawingChecker.ClearHighlight()
            Dim finding As Finding = TryCast(Grid.CurrentRow.Tag, Finding)
            If finding Is Nothing OrElse finding.CurItem Is Nothing OrElse finding.CurItem.Native Is Nothing Then Throw New Exception("This result has no current NX object. A missing item can only be located in the reference inventory.")
            If DrawingChecker.Session.Parts.Display.Tag <> OwnerPart.Tag Then Throw New Exception("Display the current inspected drawing before highlighting.")
            Dim displayable As DisplayableObject = TryCast(finding.CurItem.Native, DisplayableObject)
            If displayable Is Nothing Then Throw New Exception("The result refers to metadata rather than a displayable annotation.")
            DrawingChecker.ActiveHighlight = displayable : displayable.Highlight()
            Status.Text = "Highlighted current object; sheet=" & finding.CurItem.Sheet & "; view=" & finding.CurItem.ViewName & "; native origin=" & Inspector.Location(finding.CurItem)
        Catch ex As Exception
            F.MessageBox.Show(ex.Message, "Locate/highlight", F.MessageBoxButtons.OK, F.MessageBoxIcon.Information)
        End Try
    End Sub
    Private Sub Export(sender As Object, args As EventArgs)
        If Busy OrElse RunFolder = "" Then Return
        Try
            Busy = True : SetBusy(True) : ExportButton.Enabled = False
            Progress("Writing Excel workbook...")
            Reports.Excel(RunFolder, RefData, CurData, Rows, Options, ExtractionOnly, CurrentImage, ReferenceImageAtRun)
            Status.Text = "Saved Excel report: " & RunFolder
            F.MessageBox.Show(Status.Text, "Drawing checker")
        Catch ex As Exception
            DrawingChecker.Log("ERROR Excel: " & ex.ToString())
            F.MessageBox.Show("Excel export failed: " & ex.Message & vbCrLf & "CSV reports remain in " & RunFolder, "Drawing checker", F.MessageBoxButtons.OK, F.MessageBoxIcon.Warning)
        Finally
            Busy = False : SetBusy(False) : ExportButton.Enabled = True
            If RunFolder <> "" Then File.WriteAllText(Path.Combine(RunFolder, "Run_Log.txt"), String.Join(vbCrLf, DrawingChecker.Diagnostics.ToArray()), New UTF8Encoding(False))
        End Try
    End Sub
    Private Sub CaptureCurrent(filename As String)
        Dim builder As ImageExportBuilder = Nothing
        Try
            builder = OwnerPart.Views.CreateImageExportBuilder()
            builder.FileName = filename : builder.FileFormat = ImageExportBuilder.FileFormats.Png
            builder.RegionMode = False : builder.BackgroundOption = ImageExportBuilder.BackgroundOptions.Original
            builder.Commit()
        Finally
            If builder IsNot Nothing Then builder.Destroy()
        End Try
        If Not File.Exists(filename) AndAlso File.Exists(filename & ".png") Then File.Move(filename & ".png", filename)
        If Not File.Exists(filename) Then Throw New IOException("NX did not create the graphics image.")
    End Sub
    Private Sub ClosingForm(sender As Object, args As F.FormClosingEventArgs)
        If Busy Then
            CancelRequested = True : args.Cancel = True
        Else
            DrawingChecker.ClearHighlight()
        End If
    End Sub
End Class

' -------------------- CSV and Excel reporting --------------------
Public Class Reports
    Public Shared Function StatusColour(status As String) As D.Color
        Select Case status
            Case "FAIL", "MISSING", "ERROR" : Return D.Color.MistyRose
            Case "WARNING", "UNVERIFIED" : Return D.Color.LemonChiffon
            Case "EXTRA" : Return D.Color.LightBlue
            Case "PASS" : Return D.Color.Honeydew
            Case Else : Return D.Color.White
        End Select
    End Function
    Private Shared Function Count(rows As List(Of Finding), status As String) As Integer
        Dim value As Integer = 0
        For Each row As Finding In rows
            If row.Status = status Then value += 1
        Next
        Return value
    End Function
    Private Shared Function Summary(reference As Snapshot, current As Snapshot, rows As List(Of Finding), options As Settings, extractionOnly As Boolean) As List(Of String())
        Dim data As New List(Of String())()
        data.Add(New String() {"Field", "Value"})
        data.Add(New String() {"Overall", Inspector.Overall(rows, extractionOnly)})
        data.Add(New String() {"Current drawing", current.PartPath})
        data.Add(New String() {"Reference drawing", reference.PartPath})
        data.Add(New String() {"Part number (mapped attribute)", Mapped(current, "PartNumber")})
        data.Add(New String() {"Revision (mapped attribute)", Mapped(current, "Revision")})
        data.Add(New String() {"Check date (local)", DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss zzz")})
        data.Add(New String() {"Version", DrawingChecker.Version})
        data.Add(New String() {"Mode", If(extractionOnly, "Extraction only", options.Mode)})
        data.Add(New String() {"Checks (field/rule rows, not object count)", rows.Count.ToString()})
        For Each status As String In New String() {"PASS", "FAIL", "MISSING", "EXTRA", "WARNING", "UNVERIFIED", "ERROR"}
            data.Add(New String() {status, Count(rows, status).ToString()})
        Next
        data.Add(New String() {"Current extracted objects", current.Items.Count.ToString()})
        data.Add(New String() {"Reference extracted objects", reference.Items.Count.ToString()})
        data.Add(New String() {"Numeric equality epsilon (native units)", options.NumericEpsilon.ToString("R", CultureInfo.InvariantCulture)})
        data.Add(New String() {"Layout warning threshold (mm)", options.LayoutMm.ToString(CultureInfo.InvariantCulture)})
        data.Add(New String() {"Journal ID trust enabled", options.TrustJournalIds.ToString()})
        data.Add(New String() {"Categories", String.Join("; ", New List(Of String)(options.Categories).ToArray())})
        data.Add(New String() {"Current sheets", If(options.CurrentAllSheets, "All", String.Join(";", New List(Of String)(options.CurrentSheets).ToArray()))})
        data.Add(New String() {"Reference sheets", If(options.ReferenceAllSheets, "All", String.Join(";", New List(Of String)(options.ReferenceSheets).ToArray()))})
        data.Add(New String() {"Current views", If(options.CurrentAllViews, "All", String.Join(";", New List(Of String)(options.CurrentViews).ToArray()))})
        data.Add(New String() {"Reference views", If(options.ReferenceAllViews, "All", String.Join(";", New List(Of String)(options.ReferenceViews).ToArray()))})
        data.Add(New String() {"Company profile", options.ProfilePath})
        data.Add(New String() {"Approved vocabulary", options.DictionaryPath})
        data.Add(New String() {"Interpretation", "This report is a staged automated comparison, not engineering-release approval. PASS applies only to the named extracted field and assumed correspondence."})
        Return data
    End Function
    Private Shared Function Mapped(data As Snapshot, name As String) As String
        For Each i As Item In data.Items
            If i.Category = "Title block" AndAlso i.Name = name Then Return If(i.State = "OK", i.RawText, "UNVERIFIED")
        Next
        Return "UNVERIFIED - no configured mapping"
    End Function
    Private Shared Function Detail(rows As List(Of Finding), criticalOnly As Boolean) As List(Of String())
        Dim data As New List(Of String())()
        data.Add(New String() {"ID", "Category", "Check", "Reference", "Current", "Result", "Match method", "Reference sheet", "Current sheet", "Reference view", "Current view", "Reference object ID", "Current object ID", "Reference native origin", "Current native origin", "Diagnostic"})
        For Each row As Finding In rows
            If criticalOnly AndAlso row.Status = "PASS" Then Continue For
            data.Add(New String() {row.ID, row.Category, row.Rule, row.Reference, row.CurrentValue, row.Status, row.MatchMethod, If(row.RefItem Is Nothing, "", row.RefItem.Sheet), If(row.CurItem Is Nothing, "", row.CurItem.Sheet), If(row.RefItem Is Nothing, "", row.RefItem.ViewName), If(row.CurItem Is Nothing, "", row.CurItem.ViewName), If(row.RefItem Is Nothing, "", row.RefItem.Identity()), If(row.CurItem Is Nothing, "", row.CurItem.Identity()), Inspector.Location(row.RefItem), Inspector.Location(row.CurItem), row.Message})
        Next
        Return data
    End Function
    Private Shared Function Inventory(snapshot As Snapshot) As List(Of String())
        Dim data As New List(Of String())()
        data.Add(New String() {"Side", "Identity", "Category", "NX type", "Name", "CHECK_ID", "Journal ID", "Sheet", "View", "Native origin", "Part unit", "State", "Problem", "Field", "Value"})
        For Each i As Item In snapshot.Items
            Dim fields As New SortedDictionary(Of String, String)(i.Fields, StringComparer.Ordinal)
            fields("Raw displayed text") = i.RawText
            For Each pair As KeyValuePair(Of String, Double) In i.Numbers
                fields(pair.Key) = pair.Value.ToString("R", CultureInfo.InvariantCulture)
            Next
            For Each pair As KeyValuePair(Of String, String) In fields
                data.Add(New String() {snapshot.Label, i.Identity(), i.Category, i.Kind, i.Name, i.CheckId, i.JournalId, i.Sheet, i.ViewName, Inspector.Location(i), i.Unit, i.State, i.Problem, pair.Key, pair.Value})
            Next
        Next
        Return data
    End Function
    Public Shared Sub WriteCsv(folder As String, reference As Snapshot, current As Snapshot, rows As List(Of Finding), options As Settings, extractionOnly As Boolean)
        Csv(Path.Combine(folder, "Inspection_Summary.csv"), Summary(reference, current, rows, options, extractionOnly))
        Csv(Path.Combine(folder, "Detailed_Comparison.csv"), Detail(rows, False))
        Csv(Path.Combine(folder, "Missing_and_Modified.csv"), Detail(rows, True))
        Csv(Path.Combine(folder, "Current_Inventory.csv"), Inventory(current))
        Csv(Path.Combine(folder, "Reference_Inventory.csv"), Inventory(reference))
        File.WriteAllText(Path.Combine(folder, "Effective_Profile.ini"), If(options.ProfilePath = "", "# No company profile supplied.", File.ReadAllText(options.ProfilePath, Encoding.UTF8)), New UTF8Encoding(False))
        If options.DictionaryPath <> "" Then File.Copy(options.DictionaryPath, Path.Combine(folder, "Effective_Vocabulary.txt"), False)
    End Sub
    Private Shared Sub Csv(filename As String, rows As List(Of String()))
        Using writer As New StreamWriter(filename, False, New UTF8Encoding(True))
            For Each row() As String In rows
                Dim fields As New List(Of String)()
                For Each text As String In row
                    Dim value As String = If(text, "")
                    ' Neutralize spreadsheet formula injection; originals remain in in-memory records.
                    If Regex.IsMatch(value, "^\s*[=+@-]") Then value = "'" & value
                    fields.Add("""" & value.Replace("""", """""") & """")
                Next
                writer.WriteLine(String.Join(",", fields.ToArray()))
            Next
        End Using
    End Sub
    Public Shared Sub Excel(folder As String, reference As Snapshot, current As Snapshot, rows As List(Of Finding), options As Settings, extractionOnly As Boolean, currentImage As String, referenceImage As String)
        Dim com As New ComScope(), app As Object = Nothing, book As Object = Nothing
        Try
            app = com.Keep(CreateObject("Excel.Application"))
            app.Visible = False : app.DisplayAlerts = False
            Dim books As Object = com.Keep(app.Workbooks)
            book = com.Keep(books.Add(-4167)) ' one worksheet, private Excel instance
            Dim sheets As Object = com.Keep(book.Worksheets)
            Dim summarySheet As Object = com.Keep(sheets.Item(1)) : summarySheet.Name = "Inspection Summary"
            WriteSheet(summarySheet, Summary(reference, current, rows, options, extractionOnly), com, 0)
            Dim detailSheet As Object = AddSheet(sheets, "Detailed Comparison", com)
            WriteSheet(detailSheet, Detail(rows, False), com, 6)
            Dim issuesSheet As Object = AddSheet(sheets, "Missing and Modified Items", com)
            WriteSheet(issuesSheet, Detail(rows, True), com, 6)
            Dim images As Object = AddSheet(sheets, "Drawing Images", com)
            WriteSheet(images, New List(Of String()) From {New String() {"Evidence", "Description"}, New String() {"Current image", "Current displayed graphics area only; may show one sheet, not the complete inspection scope."}, New String() {"Reference image", "Optional user-supplied PNG/JPEG. Provenance and sheet identity must be checked manually."}, New String() {"Image comparison", "UNVERIFIED: no image registration, geometry comparison or reliable difference-marker projection implemented."}}, com, 0)
            AddImage(images, currentImage, "Current active graphics", "A7", com)
            AddImage(images, referenceImage, "Reference supplied by user", "J7", com)
            Dim validation As Object = AddSheet(sheets, "Validation Log", com)
            Dim log As New List(Of String()) From {New String() {"Level", "Diagnostic"}}
            log.Add(New String() {"VERSION", DrawingChecker.Version})
            log.Add(New String() {"VALIDATION", "Not compiled or executed with NX 2306 / Excel here. Run VALIDATION.md on copied engineering drawings."})
            For Each row As Finding In rows
                If row.Status = "UNVERIFIED" OrElse row.Status = "ERROR" OrElse row.Status = "WARNING" Then log.Add(New String() {row.Status, row.ID & " " & row.Rule & ": " & row.Message})
            Next
            For Each line As String In DrawingChecker.Diagnostics
                log.Add(New String() {"RUN", line})
            Next
            WriteSheet(validation, log, com, 0)
            Dim rawCurrent As Object = AddSheet(sheets, "Current Inventory", com)
            WriteSheet(rawCurrent, Inventory(current), com, 0)
            Dim rawReference As Object = AddSheet(sheets, "Reference Inventory", com)
            WriteSheet(rawReference, Inventory(reference), com, 0)
            summarySheet.Activate()
            Dim filename As String = Path.Combine(folder, "Drawing_Inspection.xlsx")
            If File.Exists(filename) Then filename = Path.Combine(folder, "Drawing_Inspection_" & DateTime.Now.ToString("HHmmss") & "_" & Guid.NewGuid().ToString("N").Substring(0, 6) & ".xlsx")
            book.SaveAs(filename, 51)
            DrawingChecker.Log("Workbook saved: " & filename)
        Finally
            If book IsNot Nothing Then
                Try
                    book.Close(False)
                Catch ex As Exception
                    DrawingChecker.Log("WARNING Excel workbook cleanup: " & ex.Message)
                End Try
            End If
            If app IsNot Nothing Then
                Try
                    app.Quit()
                Catch ex As Exception
                    DrawingChecker.Log("WARNING Excel process cleanup: " & ex.Message)
                End Try
            End If
            com.ReleaseAll()
        End Try
    End Sub
    Private Shared Function AddSheet(sheets As Object, name As String, com As ComScope) As Object
        Dim last As Object = com.Keep(sheets.Item(sheets.Count))
        Dim sheet As Object = com.Keep(sheets.Add(Type.Missing, last))
        sheet.Name = name : Return sheet
    End Function
    Private Shared Sub WriteSheet(sheet As Object, rows As List(Of String()), com As ComScope, statusColumn As Integer)
        If rows.Count = 0 Then Return
        If rows.Count > 1048576 Then Throw New Exception("Excel row limit exceeded; use the full CSV inventory.")
        Dim columns As Integer = rows(0).Length
        Dim values(rows.Count - 1, columns - 1) As Object
        For r As Integer = 0 To rows.Count - 1
            For c As Integer = 0 To columns - 1
                Dim value As String = If(rows(r)(c), "")
                If value.Length > 32760 Then value = value.Substring(0, 32690) & " [TRUNCATED FOR EXCEL - full value in CSV]"
                values(r, c) = value
            Next
        Next
        Dim endColumn As String = ColumnName(columns)
        Dim range As Object = com.Keep(sheet.Range("A1:" & endColumn & rows.Count.ToString()))
        range.NumberFormat = "@" : range.Value2 = values : range.WrapText = True : range.VerticalAlignment = -4160
        Dim font As Object = com.Keep(range.Font) : font.Name = "Calibri" : font.Size = 10
        Dim header As Object = com.Keep(sheet.Range("A1:" & endColumn & "1"))
        Dim hf As Object = com.Keep(header.Font) : hf.Bold = True : hf.Color = D.ColorTranslator.ToOle(D.Color.White)
        Dim hi As Object = com.Keep(header.Interior) : hi.Color = D.ColorTranslator.ToOle(D.Color.FromArgb(28, 52, 74))
        header.RowHeight = 30
        range.AutoFilter()
        For c As Integer = 1 To columns
            Dim col As Object = com.Keep(sheet.Range(ColumnName(c) & ":" & ColumnName(c)))
            col.ColumnWidth = If(columns = 2, If(c = 1, 38, 105), If(c = 4 OrElse c = 5 OrElse c = 16, 55, 25))
        Next
        Dim body As Object = com.Keep(sheet.Range("A2:" & endColumn & Math.Max(2, rows.Count).ToString()))
        body.RowHeight = If(columns = 2, 40, 58)
        If statusColumn > 0 Then
            For r As Integer = 1 To rows.Count - 1
                Dim line As Object = com.Keep(sheet.Range("A" & (r + 1).ToString() & ":" & endColumn & (r + 1).ToString()))
                Dim fill As Object = com.Keep(line.Interior)
                fill.Color = D.ColorTranslator.ToOle(StatusColour(rows(r)(statusColumn - 1)))
            Next
        End If
        sheet.Activate()
        Dim excelApp As Object = com.Keep(sheet.Application)
        Dim window As Object = com.Keep(excelApp.ActiveWindow)
        window.SplitRow = 1 : window.FreezePanes = True
    End Sub
    Private Shared Function ColumnName(index As Integer) As String
        Dim text As String = ""
        While index > 0
            index -= 1 : text = ChrW(65 + index Mod 26) & text : index = index \ 26
        End While
        Return text
    End Function
    Private Shared Sub AddImage(sheet As Object, filename As String, caption As String, anchorName As String, com As ComScope)
        Dim anchor As Object = com.Keep(sheet.Range(anchorName))
        anchor.Value2 = caption
        If filename = "" Then Return
        If Not File.Exists(filename) Then Throw New FileNotFoundException("Selected image no longer exists.", filename)
        Using test As D.Image = D.Image.FromFile(filename)
            If test.Width < 10 OrElse test.Height < 10 Then Throw New IOException("Invalid image: " & filename)
        End Using
        Dim shapes As Object = com.Keep(sheet.Shapes)
        Dim picture As Object = com.Keep(shapes.AddPicture(filename, 0, -1, CDbl(anchor.Left), CDbl(anchor.Top) + 25, -1, -1))
        picture.LockAspectRatio = -1 : picture.Width = 510.0
        If CDbl(picture.Height) > 650.0 Then picture.Height = 650.0
    End Sub
End Class
Public Class ComScope
    Private Objects As New List(Of Object)()
    Public Function Keep(value As Object) As Object
        For Each existing As Object In Objects
            If Object.ReferenceEquals(existing, value) Then Return value
        Next
        Objects.Add(value) : Return value
    End Function
    Public Sub ReleaseAll()
        For index As Integer = Objects.Count - 1 To 0 Step -1
            Try
                If Marshal.IsComObject(Objects(index)) Then Marshal.FinalReleaseComObject(Objects(index))
            Catch
                ' Cleanup must not hide the export failure.
            End Try
        Next
        Objects.Clear()
    End Sub
End Class

' -------------------- Pure engine regression tests; execute before any NX read --------------------
Public Class EngineTests
    Private Shared Sub Require(condition As Boolean, message As String)
        If Not condition Then Throw New Exception("Comparison engine self-test failed: " & message)
    End Sub
    Private Shared Function DimItem(key As String, id As String, tolerance As Double) As Item
        Dim i As New Item With {.Key = key, .JournalId = key, .CheckId = id, .Category = "Dimensions", .Kind = "DiameterDimension", .Unit = "Millimeters", .RawText = "<o>25"}
        i.Numbers("Computed nominal") = 25.0 : i.Numbers("Upper API tolerance") = tolerance
        Return i
    End Function
    Public Shared Sub Run()
        Dim settings As New Settings()
        Dim r As New Snapshot(), c As New Snapshot()
        r.Items.Add(DimItem("r1", "HOLE_1", 0.05)) : c.Items.Add(DimItem("c1", "HOLE_1", 0.1))
        Dim rows As List(Of Finding) = Inspector.Compare(r, c, settings)
        Dim found As Boolean = False
        For Each row As Finding In rows
            If row.Rule = "Upper API tolerance" AndAlso row.Status = "FAIL" Then found = True
        Next
        Require(found, "tolerance growth must FAIL")
        Dim oldId As String = ""
        For Each row As Finding In rows
            If row.Rule = "Upper API tolerance" Then oldId = row.ID
        Next
        c.Items(0).Numbers("Upper API tolerance") = 0.2
        rows = Inspector.Compare(r, c, settings)
        For Each row As Finding In rows
            If row.Rule = "Upper API tolerance" Then Require(row.ID = oldId, "result ID changed with field value")
        Next
        r = New Snapshot() : c = New Snapshot()
        r.Items.Add(DimItem("r1", "", 0.05)) : r.Items.Add(DimItem("r2", "", 0.05))
        c.Items.Add(DimItem("c1", "", 0.05)) : c.Items.Add(DimItem("c2", "", 0.05))
        rows = Inspector.Compare(r, c, settings)
        For Each row As Finding In rows
            Require(row.Status <> "PASS", "duplicate content must remain ambiguous")
        Next
        r = New Snapshot() : c = New Snapshot()
        Dim rf As New Item With {.Key = "rF", .CheckId = "FCF_1", .Category = "GD&T", .Kind = "DraftingFcf"}
        Dim cf As New Item With {.Key = "cF", .CheckId = "FCF_1", .Category = "GD&T", .Kind = "DraftingFcf"}
        rf.Put("Frame[0].Secondary.Letter", "B") : cf.Put("Frame[0].Secondary.Letter", "C")
        r.Items.Add(rf) : c.Items.Add(cf)
        rows = Inspector.Compare(r, c, settings)
        Require(rows.Count = 1 AndAlso rows(0).Status = "FAIL", "datum-role change must FAIL")
        r = New Snapshot() : c = New Snapshot()
        r.Items.Add(DimItem("r1", "HOLE_1", 0.05)) : c.Items.Add(DimItem("c1", "HOLE_1", 0.05))
        c.Items(0).Unit = "Inches"
        rows = Inspector.Compare(r, c, settings)
        For Each row As Finding In rows
            If row.Rule = "Computed nominal" Then Require(row.Status = "UNVERIFIED", "mixed units must not PASS")
        Next
        r.Items(0).RawText = "SAE  1045" : c.Items(0).RawText = "SAE 1045"
        settings.Mode = "Engineering" : rows = Inspector.Compare(r, c, settings)
        For Each row As Finding In rows
            If row.Rule = "Displayed text" Then Require(row.Status = "WARNING", "engineering whitespace comparison")
        Next
        settings.Mode = "Exact" : rows = Inspector.Compare(r, c, settings)
        For Each row As Finding In rows
            If row.Rule = "Displayed text" Then Require(row.Status = "FAIL", "exact whitespace comparison")
        Next
        Require(Inspector.Normalize("SAE 1045") <> Inspector.Normalize("SAE 1050"), "normalization removed technical numbers")
        Require(Inspector.Normalize("A-B") <> Inspector.Normalize("A B"), "normalization merged common and sequential datum expressions")
        Require(Inspector.Normalize("25 +/- 0.05") <> Inspector.Normalize("25 +/- 0.10"), "normalization removed tolerance")
        Dim one As New List(Of Finding) From {Inspector.Make("UNVERIFIED", "Coverage", "test", Nothing, Nothing, "", "", "")}
        Require(Inspector.Overall(one, False).IndexOf("PASS", StringComparison.Ordinal) >= 0 AndAlso Inspector.Overall(one, False).StartsWith("INCOMPLETE", StringComparison.Ordinal), "unverified status must prevent overall PASS")
        one.Add(Inspector.Make("FAIL", "Dimensions", "test", Nothing, Nothing, "", "", ""))
        Require(Inspector.Overall(one, False) = "FAIL / INCOMPLETE", "mixed failure and incomplete status")
        Require(Inspector.Overall(one, True) = "EXTRACTED - NOT INSPECTED", "extraction must not imply inspection")
    End Sub
End Class
