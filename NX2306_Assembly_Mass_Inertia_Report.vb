' NX2306_Assembly_Mass_Inertia_Report.vb
' NX 2306 / VB.NET journal. No model geometry, visibility, WCS or assembly position is changed.
' NOTE: Requires Windows NX journal runtime, and Excel for .xlsx (CSV fallback provided).
Option Strict Off
Option Explicit On
Imports System
Imports System.IO
Imports System.Text
Imports System.Globalization
Imports System.Collections.Generic
Imports System.Reflection
Imports System.Runtime.InteropServices
Imports System.Windows.Forms
Imports NXOpen
Imports NXOpen.UF
Imports NXOpen.Utilities
Imports NXOpen.Assemblies

Module NX2306_Assembly_Mass_Inertia_Report
    Private sess As Session
    Private uf As UFSession
    Private lw As ListingWindow
    Private top As Part
    Private warnings As New List(Of String)
    Private records As New Dictionary(Of String, CompRec)(StringComparer.Ordinal)
    Private missingMap As New Dictionary(Of String, HashSet(Of String))(StringComparer.Ordinal)
    Private visited As New HashSet(Of String)(StringComparer.Ordinal)
    Private isIncomplete As Boolean = False
    Private counted As Integer = 0
    Private isVisibleOnly As Boolean = False
    Private densityFallback As Double = 0.0   ' kg/m^3; zero = not authorized
    Private reportFile As String = ""
    Private makeExcel As Boolean = True
    Private makeTable As Boolean = False
    Private totalM As Double = 0.0
    Private combinedCG(2) As Double
    Private combinedI(2, 2) As Double
    Private principal(2) As Double
    Private capturePath As String = ""

    Private Class CompRec
        Public Path As String
        Public PartName As String
        Public Mass As Double
        Public S(2) As Double
        Public CG(2) As Double
        Public Cnt As Integer
        Public ExpectedSolids As Integer
        Public Status As String = "OK"
    End Class

    Private Class BodyRec
        Public M As Double
        Public C(2) As Double
        Public I(2, 2) As Double  ' central tensor, kg-mm^2; OFF-DIAGONAL = -integral(x*y*dm)
        Public CompKey As String
        Public PartName As String
    End Class
    Private bodies As New List(Of BodyRec)

    Public Sub Main()
        Try
            sess = Session.GetSession()
            uf = UFSession.GetUFSession()
            lw = sess.ListingWindow
            lw.Open()
            top = sess.Parts.Display
            If top Is Nothing Then
                MessageBox.Show("Open an assembly as the displayed part.", "Mass properties")
                Return
            End If
            If top.ComponentAssembly.RootComponent Is Nothing Then
                MessageBox.Show("The displayed part is not a top-level assembly.", "Mass properties")
                Return
            End If
            If sess.Parts.Work Is Nothing OrElse sess.Parts.Work.Tag <> top.Tag Then
                MessageBox.Show("For an unambiguous assembly absolute/WCS reference, make the displayed top-level assembly the WORK part before running this journal. No work-part switch is performed automatically.", "Mass properties")
                Return
            End If
            If Not OptionsDialog() Then Return
            If Not top.IsFullyLoaded Then Warn("TOP ASSEMBLY PART NOT FULLY LOADED.", True)
            Say("RUN: " & DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
            Say("Display assembly: " & top.FullPath)
            Say("Scope: " & If(isVisibleOnly, "Visible occurrences only", "All active reference-set occurrence bodies"))
            Say("Density fallback (kg/m^3): " & If(densityFallback > 0, densityFallback.ToString("G", CultureInfo.InvariantCulture), "NONE"))
            Say("Units: kg, mm, kg-mm^2. CG in assembly ABS, tensor at assembly CG, ABS-parallel axes.")
            Say("Tensor convention: off-diagonal Jxy=-integral(x*y dm); reported product Ixy=+integral(x*y dm).")

            Traverse(top.ComponentAssembly.RootComponent, "", True)
            Dim sourceTags As New HashSet(Of String)(StringComparer.Ordinal)
            Dim tag As Tag = Tag.Null
            Do
                uf.Obj.CycleObjsInPart(top.Tag, UFConstants.UF_solid_type, tag)
                If tag = Tag.Null Then Exit Do
                Dim typ As Integer = 0, subtyp As Integer = 0
                uf.Obj.AskTypeAndSubtype(tag, typ, subtyp)
                If subtyp <> UFConstants.UF_solid_body_subtype Then Continue Do
                If Not sourceTags.Add(tag.ToString()) Then Continue Do
                Dim body As Body = TryCast(NXObjectManager.Get(tag), Body)
                If body Is Nothing OrElse Not body.IsSolidBody Then Continue Do
                If isVisibleOnly AndAlso (body.IsBlanked OrElse HiddenByParent(body.OwningComponent)) Then Continue Do
                Try
                    ProcessBody(body)
                Catch ex As Exception
                    Dim label As String = "Body tag " & tag.ToString()
                    Try
                        If body.OwningComponent IsNot Nothing Then label = MakePath(body.OwningComponent)
                    Catch
                    End Try
                    Warn("MEASUREMENT FAILED: " & label & ": " & ex.Message, True)
                    Dim key As String = If(body.OwningComponent Is Nothing, "[TOP_ASSEMBLY_BODIES]", MakePath(body.OwningComponent))
                    MarkRow(key, If(body.OwningPart Is Nothing, "Unknown", body.OwningPart.Leaf), "FAILED")
                End Try
            Loop
            ' An unsuppressed, loaded component containing modeled solids but no occurrence
            ' geometry is not silently reported as zero mass. Such parts may be excluded by
            ' a reference set or unavailable in the current display context.
            For Each key As String In visited
                Dim r As CompRec = records(key)
                If r.Cnt = 0 AndAlso r.Status = "PENDING_SOLID" Then
                    If isVisibleOnly Then
                        Warn("NO VISIBLE OCCURRENCE BODIES (filtered/check reference set): " & key, False)
                        r.Status = "NOT VISIBLE/FILTERED"
                    Else
                        Warn("NO MEASURABLE OCCURRENCE BODIES: " & key & " (reference set/load/representation).", True)
                        r.Status = "NO GEOMETRY"
                    End If
                ElseIf r.Status = "PENDING_SOLID" Then
                    r.Status = "OK"
                End If
                If Not isVisibleOnly AndAlso r.Cnt < r.ExpectedSolids AndAlso
                   (r.Status = "OK" OrElse r.Status = "FALLBACK DENSITY") Then
                    Warn("FEWER BODY OCCURRENCES THAN PROTOTYPE SOLIDS: " & key &
                         " measured=" & r.Cnt.ToString() & " expected=" & r.ExpectedSolids.ToString() &
                         " (check reference sets, unloaded bodies, and representation).", True)
                    r.Status = "PARTIAL OCCURRENCE"
                End If
            Next
            Aggregate()
            WriteListing()

            If makeExcel Then
                capturePath = Path.Combine(Path.GetDirectoryName(reportFile), Path.GetFileNameWithoutExtension(reportFile) & "_NX_View.png")
                Try
                    CaptureImage(capturePath)
                Catch ex As Exception
                    capturePath = ""
                    Warn("Graphics capture failed: " & ex.Message, False)
                    Say("The workbook will be exported without a picture.")
                End Try
                Try
                    ExportXlsx(reportFile)
                    Say("XLSX saved: " & reportFile)
                Catch ex As Exception
                    Warn("Excel COM unavailable/export failed: " & ex.Message, False)
                    ExportCsv(reportFile)
                    Say("CSV files saved adjacent to: " & reportFile)
                End Try
            End If
            If makeTable Then
                Try
                    CreateNXTable()
                Catch ex As Exception
                    Warn("Native NX tabular note not created: " & ex.Message, False)
                End Try
            End If
            If warnings.Count > 0 Then
                Say("WARNINGS:")
                For Each s As String In warnings
                    Say("  - " & s)
                Next
            End If
            MessageBox.Show("Mass properties completed: " & If(isIncomplete, "INCOMPLETE - see warnings", "MEASURED") & vbCrLf &
                            "Measured bodies: " & counted.ToString() & vbCrLf &
                            "Total mass: " & F(totalM) & " kg" & vbCrLf &
                            If(makeExcel, "Output: " & reportFile & " (or CSV fallback)", "Excel export not selected"),
                            "Assembly mass properties")
        Catch ex As Exception
            If lw IsNot Nothing Then Say("FATAL: " & ex.ToString())
            MessageBox.Show(ex.Message, "NX mass report error")
        End Try
    End Sub

    Private Function OptionsDialog() As Boolean
        Using frm As New Form()
            frm.Text = "NX 2306 - Assembly CG and MI Report"
            frm.Width = 560 : frm.Height = 330
            frm.StartPosition = FormStartPosition.CenterScreen
            frm.FormBorderStyle = FormBorderStyle.FixedDialog
            frm.MaximizeBox = False : frm.MinimizeBox = False
            Dim labelScope As New Label() With {.Text = "Component scope", .Left = 18, .Top = 20, .Width = 165}
            Dim scope As New ComboBox() With {.Left = 194, .Top = 16, .Width = 322, .DropDownStyle = ComboBoxStyle.DropDownList}
            scope.Items.AddRange(New Object() {"Entire active assembly", "Visible solid bodies only"}) : scope.SelectedIndex = 0
            Dim labelDen As New Label() With {.Text = "Fallback density", .Left = 18, .Top = 66, .Width = 165}
            Dim inp As New TextBox() With {.Left = 194, .Top = 62, .Width = 143, .Text = ""}
            Dim u As New ComboBox() With {.Left = 350, .Top = 62, .Width = 166, .DropDownStyle = ComboBoxStyle.DropDownList}
            u.Items.AddRange(New Object() {"kg/m^3", "g/cm^3"}) : u.SelectedIndex = 0
            Dim advice As New Label() With {.Left = 18, .Top = 94, .Width = 506, .Height = 43,
                .Text = "Leave density empty to EXCLUDE bodies without material. Enter a positive fallback explicitly; it never changes the NX body."}
            Dim excelOpt As New CheckBox() With {.Left = 18, .Top = 148, .Width = 450, .Text = "Export Excel .xlsx (requires Microsoft Excel)", .Checked = True}
            Dim tabOpt As New CheckBox() With {.Left = 18, .Top = 182, .Width = 480, .Text = "Create optional native NX tabular note", .Checked = False}
            Dim ok As New Button() With {.Text = "Continue", .Left = 328, .Top = 231, .Width = 92, .DialogResult = DialogResult.OK}
            Dim cancel As New Button() With {.Text = "Cancel", .Left = 429, .Top = 231, .Width = 92, .DialogResult = DialogResult.Cancel}
            frm.Controls.AddRange(New Control() {labelScope, scope, labelDen, inp, u, advice, excelOpt, tabOpt, ok, cancel})
            frm.AcceptButton = ok : frm.CancelButton = cancel
            If frm.ShowDialog() <> DialogResult.OK Then Return False
            isVisibleOnly = (scope.SelectedIndex = 1)
            densityFallback = 0.0
            If inp.Text.Trim() <> "" Then
                Dim entered As Double
                If Not Double.TryParse(inp.Text.Trim(), NumberStyles.Float, CultureInfo.CurrentCulture, entered) AndAlso
                   Not Double.TryParse(inp.Text.Trim(), NumberStyles.Float, CultureInfo.InvariantCulture, entered) Then
                    MessageBox.Show("Fallback density is not a valid number; operation cancelled.") : Return False
                End If
                If Double.IsNaN(entered) OrElse Double.IsInfinity(entered) OrElse entered <= 0 Then
                    MessageBox.Show("Fallback density must be positive; operation cancelled.") : Return False
                End If
                densityFallback = entered * If(u.SelectedIndex = 1, 1000.0, 1.0)
            End If
            makeExcel = excelOpt.Checked : makeTable = tabOpt.Checked
        End Using
        If makeExcel Then
            Using sf As New SaveFileDialog()
                sf.Title = "Save assembly mass-properties report"
                sf.Filter = "Excel workbook (*.xlsx)|*.xlsx"
                sf.AddExtension = True : sf.DefaultExt = "xlsx"
                sf.FileName = Path.GetFileNameWithoutExtension(top.Leaf) & "_Mass_Inertia.xlsx"
                sf.OverwritePrompt = True
                If sf.ShowDialog() <> DialogResult.OK Then Return False
                reportFile = sf.FileName
            End Using
        End If
        Return True
    End Function

    Private Sub Traverse(c As Component, prefix As String, isRoot As Boolean)
        If c Is Nothing Then Return
        If Not isRoot AndAlso c.IsSuppressed Then
            Warn("SUPPRESSED COMPONENT excluded: " & If(prefix = "", c.DisplayName, prefix & "/" & c.DisplayName), False)
            Return
        End If
        Dim occurrenceName As String = c.DisplayName & "{tag=" & c.Tag.ToString() & "}"
        Dim key As String = If(isRoot, "[ROOT]", If(prefix = "", occurrenceName, prefix & "/" & occurrenceName))
        If Not isRoot Then
            visited.Add(key)
            Dim pp As Part = TryCast(c.Prototype, Part)
            Dim partName As String = If(pp Is Nothing, "<unloaded>", pp.Leaf)
            Dim r As CompRec = EnsureRow(key, partName)
            If pp Is Nothing Then
                r.Status = "UNLOADED"
                Warn("UNLOADED COMPONENT: " & key, True)
                Return
            End If
            Try
                Dim ownSolids As Integer = 0
                For Each body As Body In pp.Bodies.ToArray()
                    If body.IsSolidBody Then ownSolids += 1
                Next
                r.ExpectedSolids = ownSolids
                If ownSolids > 0 Then
                    r.Status = "PENDING_SOLID"
                ElseIf c.GetChildren().Length > 0 Then
                    r.Status = "CONTAINER (SEE CHILDREN)"
                End If
                If Not pp.IsFullyLoaded Then
                    r.Status = "PARTIAL LOAD"
                    Warn("PART NOT FULLY LOADED: " & key, True)
                End If
                Dim rep As Component.RepresentationMode = c.GetComponentRepresentationMode()
                If rep = Component.RepresentationMode.RepresentationModeLightweight OrElse
                   rep = Component.RepresentationMode.RepresentationModePartial Then
                    r.Status = "LIGHTWEIGHT/PARTIAL"
                    Warn("NON-EXACT REPRESENTATION: " & key, True)
                End If
            Catch ex As Exception
                r.Status = "UNAVAILABLE"
                Warn("CANNOT INSPECT PROTOTYPE: " & key & ": " & ex.Message, True)
            End Try
        End If
        Try
            For Each child As Component In c.GetChildren()
                Traverse(child, If(isRoot, "", key), False)
            Next
        Catch ex As Exception
            Warn("CANNOT TRAVERSE CHILDREN: " & key & ": " & ex.Message, True)
        End Try
    End Sub

    Private Function HiddenByParent(c As Component) As Boolean
        Dim current As Component = c
        While current IsNot Nothing
            If current.IsBlanked Then Return True
            current = current.Parent
        End While
        Return False
    End Function

    Private Function MakePath(c As Component) As String
        If c Is Nothing Then Return "[TOP_ASSEMBLY_BODIES]"
        Dim names As New List(Of String)()
        Dim curr As Component = c
        While curr IsNot Nothing AndAlso curr.Parent IsNot Nothing
            names.Insert(0, curr.DisplayName & "{tag=" & curr.Tag.ToString() & "}")
            curr = curr.Parent
        End While
        Return String.Join("/", names.ToArray())
    End Function

    Private Function EnsureRow(key As String, pname As String) As CompRec
        If Not records.ContainsKey(key) Then records.Add(key, New CompRec() With {.Path = key, .PartName = pname})
        Return records(key)
    End Function

    Private Sub MarkRow(key As String, partName As String, status As String)
        Dim r As CompRec = EnsureRow(key, partName)
        r.Status = status
    End Sub

    ' Material detection by documented PhysicalMaterialCollection.FindBodiesWithoutMaterial.
    ' Reflection avoids depending on an unverified ByRef/ByVal declaration for its .NET output array.
    ' If the query fails, density provenance is UNKNOWN, not silently assumed.
    Private Function MaterialMissing(proto As Body, p As Part) As Nullable(Of Boolean)
        Dim k As String = p.Tag.ToString()
        If Not missingMap.ContainsKey(k) Then
            Try
                Dim coll As Object = p.MaterialManager.PhysicalMaterials
                Dim mi As MethodInfo = coll.GetType().GetMethod("FindBodiesWithoutMaterial")
                If mi Is Nothing Then Throw New Exception("FindBodiesWithoutMaterial method not in installed NXOpen.dll")
                Dim args() As Object = {Nothing}
                mi.Invoke(coll, args)
                Dim missing As New HashSet(Of String)(StringComparer.Ordinal)
                If args(0) IsNot Nothing Then
                    For Each obj As Object In CType(args(0), System.Array)
                        Dim no As TaggedObject = TryCast(obj, TaggedObject)
                        If no IsNot Nothing Then missing.Add(no.Tag.ToString())
                    Next
                End If
                missingMap.Add(k, missing)
            Catch ex As Exception
                Warn("MATERIAL ASSIGNMENT QUERY FAILED for " & p.Leaf & ": " & ex.Message, True)
                Return Nothing
            End Try
        End If
        Return missingMap(k).Contains(proto.Tag.ToString())
    End Function

    Private Function MapWcsToAbs(wcsPartUnits() As Double) As Double()
        Dim outPt(2) As Double
        uf.Csys.MapPoint(UFConstants.UF_CSYS_ROOT_WCS_COORDS, wcsPartUnits,
                         UFConstants.UF_CSYS_ROOT_COORDS, outPt)
        Return outPt
    End Function

    Private Sub ProcessBody(b As Body)
        Dim key As String = MakePath(b.OwningComponent)
        Dim pbody As Body = If(b.IsOccurrence, TryCast(b.Prototype, Body), b)
        If pbody Is Nothing Then Throw New Exception("Prototype body cannot be obtained.")
        Dim part As Part = TryCast(pbody.OwningPart, Part)
        If part Is Nothing Then Throw New Exception("Prototype owning part unavailable.")
        Dim rec As CompRec = EnsureRow(key, part.Leaf)
        Dim density As Double = pbody.Density ' documented units kg/m^3
        Dim noMaterial As Nullable(Of Boolean) = MaterialMissing(pbody, part)
        If Not noMaterial.HasValue Then
            rec.Status = "MATERIAL QUERY FAILED"
            Return
        End If
        If noMaterial.Value OrElse Not Good(density) OrElse density <= 0.0 Then
            If densityFallback <= 0 Then
                rec.Status = "MISSING DENSITY"
                Warn("EXCLUDED MISSING/INVALID DENSITY: " & key & ", body " & b.Tag.ToString(), True)
                Return
            End If
            density = densityFallback
            Warn("EXPLICIT FALLBACK DENSITY USED: " & key & ", body " & b.Tag.ToString(), False)
            If rec.Status = "OK" OrElse rec.Status = "PENDING_SOLID" Then rec.Status = "FALLBACK DENSITY"
        End If
        Dim acc(10) As Double, mp(46) As Double, stats(12) As Double
        acc(0) = 0.99
        uf.Modl.AskMassProps3d(New Tag() {b.Tag}, 1, 1, 4, 0.0, 1, acc, mp, stats)
        ' UF 4 => kg and metres. uf MassProps returns WCS centroid, centroidal inertia.
        Dim v As Double = mp(1)     ' m^3
        Dim rawM As Double = mp(2)  ' kg
        If Not Good(v) OrElse v <= 0 OrElse Not Good(rawM) OrElse rawM <= 0 Then
            Throw New Exception("Kernel provided zero or invalid mass/volume. No geometric inertia available for fallback scaling.")
        End If
        Dim rawD As Double = rawM / v ' kg/m^3
        Dim factor As Double = density / rawD
        If Not Good(factor) OrElse factor <= 0 Then Throw New Exception("Cannot scale measured mass properties to target density.")
        Dim br As New BodyRec()
        br.M = rawM * factor : br.CompKey = key : br.PartName = part.Leaf
        Dim mmPerUnit As Double = If(top.PartUnits = Part.Units.Inches, 25.4, 1.0)
        Dim wcsPos() As Double = {mp(3) * 1000.0 / mmPerUnit,
                                  mp(4) * 1000.0 / mmPerUnit,
                                  mp(5) * 1000.0 / mmPerUnit}
        Dim absPos() As Double = MapWcsToAbs(wcsPos)
        For i As Integer = 0 To 2
            br.C(i) = absPos(i) * mmPerUnit
        Next
        ' Derive WCS basis in ABS using MapPoint, without editing or moving the WCS.
        Dim origin() As Double = MapWcsToAbs(New Double() {0.0, 0.0, 0.0})
        Dim r(2, 2) As Double
        For j As Integer = 0 To 2
            Dim axis() As Double = {0.0, 0.0, 0.0}
            axis(j) = 1.0
            Dim endPt() As Double = MapWcsToAbs(axis)
            For i As Integer = 0 To 2
                r(i, j) = endPt(i) - origin(i)
            Next
        Next
        ' NX mass_props[19..21] are geometric products Pxy, Pyz, Pxz.
        Dim w(2, 2) As Double
        w(0, 0) = mp(12) : w(1, 1) = mp(13) : w(2, 2) = mp(14)
        w(0, 1) = -mp(19) : w(1, 0) = -mp(19)
        w(1, 2) = -mp(20) : w(2, 1) = -mp(20)
        w(0, 2) = -mp(21) : w(2, 0) = -mp(21)
        For i As Integer = 0 To 2
            For j As Integer = 0 To 2
                Dim s As Double = 0.0
                For a As Integer = 0 To 2
                    For d As Integer = 0 To 2
                        s += r(i, a) * w(a, d) * r(j, d)
                    Next
                Next
                br.I(i, j) = s * factor * 1000000.0  ' kg m^2 -> kg mm^2
                If Not Good(br.I(i, j)) Then Throw New Exception("Nonfinite inertia result.")
            Next
        Next
        bodies.Add(br)
        rec.Mass += br.M
        rec.Cnt += 1
        For i As Integer = 0 To 2
            rec.S(i) += br.M * br.C(i)
        Next
        If rec.Status = "PENDING_SOLID" Then rec.Status = "OK"
        counted += 1
    End Sub

    Private Sub Aggregate()
        For Each b As BodyRec In bodies
            totalM += b.M
            For i As Integer = 0 To 2
                combinedCG(i) += b.M * b.C(i)
            Next
        Next
        If totalM <= 0 Then
            Warn("No measurable positive-mass geometry found.", True)
            Return
        End If
        For i As Integer = 0 To 2
            combinedCG(i) /= totalM
        Next
        For Each b As BodyRec In bodies
            Dim d(2) As Double
            Dim dSq As Double = 0
            For i As Integer = 0 To 2
                d(i) = b.C(i) - combinedCG(i)
                dSq += d(i) * d(i)
            Next
            For i As Integer = 0 To 2
                For j As Integer = 0 To 2
                    combinedI(i, j) += b.I(i, j) + b.M *
                        (If(i = j, dSq, 0.0) - d(i) * d(j))
                Next
            Next
        Next
        principal = Eigenvalues(combinedI)
        For Each r As CompRec In records.Values
            If r.Mass > 0 Then
                For i As Integer = 0 To 2
                    r.CG(i) = r.S(i) / r.Mass
                Next
            End If
        Next
    End Sub

    ' Jacobi symmetric 3x3 diagonalization, no extra numerical library.
    Private Function Eigenvalues(a(,) As Double) As Double()
        Dim x(2, 2) As Double
        For i As Integer = 0 To 2
            For j As Integer = 0 To 2
                x(i, j) = a(i, j)
            Next
        Next
        For iter As Integer = 1 To 60
            Dim p As Integer = 0, q As Integer = 1
            For i As Integer = 0 To 2
                For j As Integer = i + 1 To 2
                    If Math.Abs(x(i, j)) > Math.Abs(x(p, q)) Then p = i : q = j
                Next
            Next
            If Math.Abs(x(p, q)) < 0.000000001 * Math.Max(1.0, Math.Abs(x(0, 0)) + Math.Abs(x(1, 1)) + Math.Abs(x(2, 2))) Then Exit For
            Dim theta As Double = 0.5 * Math.Atan2(2.0 * x(p, q), x(q, q) - x(p, p))
            Dim c As Double = Math.Cos(theta), s As Double = Math.Sin(theta)
            Dim app As Double = x(p, p), aqq As Double = x(q, q), apq As Double = x(p, q)
            x(p, p) = c * c * app - 2 * s * c * apq + s * s * aqq
            x(q, q) = s * s * app + 2 * s * c * apq + c * c * aqq
            x(p, q) = 0 : x(q, p) = 0
            For k As Integer = 0 To 2
                If k <> p AndAlso k <> q Then
                    Dim kp As Double = x(k, p), kq As Double = x(k, q)
                    x(k, p) = c * kp - s * kq : x(p, k) = x(k, p)
                    x(k, q) = s * kp + c * kq : x(q, k) = x(k, q)
                End If
            Next
        Next
        Dim ev() As Double = {x(0, 0), x(1, 1), x(2, 2)}
        Array.Sort(ev)
        Return ev
    End Function

    Private Function NumericSummary() As Double()
        Return New Double() {totalM, combinedCG(0), combinedCG(1), combinedCG(2),
                            combinedI(0, 0), combinedI(1, 1), combinedI(2, 2),
                            -combinedI(0, 1), -combinedI(0, 2), -combinedI(1, 2),
                            principal(0), principal(1), principal(2)}
    End Function

    Private Function SummaryRows() As List(Of String())
        Dim l As New List(Of String())
        l.Add(New String() {"Total mass", F(totalM), "kg"})
        l.Add(New String() {"CG X (assembly ABS)", F(combinedCG(0)), "mm"})
        l.Add(New String() {"CG Y (assembly ABS)", F(combinedCG(1)), "mm"})
        l.Add(New String() {"CG Z (assembly ABS)", F(combinedCG(2)), "mm"})
        l.Add(New String() {"Ixx @ assembly CG", F(combinedI(0, 0)), "kg-mm^2"})
        l.Add(New String() {"Iyy @ assembly CG", F(combinedI(1, 1)), "kg-mm^2"})
        l.Add(New String() {"Izz @ assembly CG", F(combinedI(2, 2)), "kg-mm^2"})
        l.Add(New String() {"Ixy (product; + integral xy)", F(-combinedI(0, 1)), "kg-mm^2"})
        l.Add(New String() {"Ixz (product; + integral xz)", F(-combinedI(0, 2)), "kg-mm^2"})
        l.Add(New String() {"Iyz (product; + integral yz)", F(-combinedI(1, 2)), "kg-mm^2"})
        l.Add(New String() {"Principal I1 (smallest)", F(principal(0)), "kg-mm^2"})
        l.Add(New String() {"Principal I2", F(principal(1)), "kg-mm^2"})
        l.Add(New String() {"Principal I3 (largest)", F(principal(2)), "kg-mm^2"})
        Return l
    End Function

    Private Sub WriteListing()
        Say("") : Say("================ ASSEMBLY MASS PROPERTIES ================")
        Say("Assembly: " & top.Leaf & "   STATUS: " & If(isIncomplete, "INCOMPLETE", "MEASURED"))
        Say("Reference: assembly absolute axes; inertia about assembly CG")
        Say("Mass and position include every measured body occurrence once.")
        Say("Metric".PadRight(38) & "Value".PadLeft(20) & "   Unit")
        Say(New String("-"c, 78))
        For Each row As String() In SummaryRows()
            Say(row(0).PadRight(38) & row(1).PadLeft(20) & "   " & row(2))
        Next
        Say(New String("-"c, 78))
        Say("Body occurrences measured: " & counted.ToString())
        Say("Component rows: " & records.Count.ToString())
    End Sub

    Private Sub CaptureImage(fileName As String)
        ' Obsolete but verified API in SC2306; current graphics window, not regenerated.
        Dim enumType As Type = GetType(UFDisp.ImageFormat)
        Dim names() As String = [Enum].GetNames(enumType)
        Dim chosen As String = ""
        For Each n As String In names
            If n.ToLowerInvariant().Contains("png") Then chosen = n : Exit For
        Next
        If chosen = "" Then Throw New Exception("PNG enum unavailable in installed UFDisp.ImageFormat: " & String.Join(",", names))
        Dim fmt As UFDisp.ImageFormat = CType([Enum].Parse(enumType, chosen), UFDisp.ImageFormat)
        ' BackgroundColor enum names vary: select a supported member; screenshot normally uses shown view.
        Dim bg As UFDisp.BackgroundColor = CType([Enum].GetValues(GetType(UFDisp.BackgroundColor)).GetValue(0), UFDisp.BackgroundColor)
        uf.Disp.CreateImage(fileName, fmt, bg)
        If Not File.Exists(fileName) Then Throw New Exception("NX did not produce a PNG image.")
    End Sub

    Private Sub ExportXlsx(fileName As String)
        Dim excel As Object = Nothing, books As Object = Nothing, book As Object = Nothing
        Dim sum As Object = Nothing, comps As Object = Nothing, warns As Object = Nothing
        Try
            excel = CreateObject("Excel.Application")
            excel.Visible = False : excel.DisplayAlerts = False
            books = excel.Workbooks
            book = books.Add()
            sum = book.Worksheets(1)
            sum.Name = "Summary"
            comps = book.Worksheets.Add(After:=sum)
            comps.Name = "Components"
            warns = book.Worksheets.Add(After:=comps)
            warns.Name = "Warnings"
            sum.Cells(1, 1).Value = "NX 2306 | Assembly Mass and Inertia"
            sum.Cells(2, 1).Value = "Assembly" : sum.Cells(2, 2).Value = top.Leaf
            sum.Cells(3, 1).Value = "Report date/time" : sum.Cells(3, 2).Value = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")
            sum.Cells(4, 1).Value = "Status" : sum.Cells(4, 2).Value = If(isIncomplete, "INCOMPLETE", "MEASURED")
            sum.Cells(5, 1).Value = "Scope" : sum.Cells(5, 2).Value = If(isVisibleOnly, "Visible bodies", "All bodies in active reference sets")
            sum.Cells(6, 1).Value = "Reference system" : sum.Cells(6, 2).Value = "Assembly absolute; inertia @ assembly CG; ABS-parallel axes"
            sum.Cells(7, 1).Value = "Products convention" : sum.Cells(7, 2).Value = "Pxy=+integral(x*y dm); tensor off-diagonal=-Pxy"
            sum.Cells(9, 1).Value = "PROPERTY" : sum.Cells(9, 2).Value = "VALUE" : sum.Cells(9, 3).Value = "UNIT"
            Dim r As Integer = 10
            Dim nums() As Double = NumericSummary()
            Dim idx As Integer = 0
            For Each item As String() In SummaryRows()
                sum.Cells(r, 1).Value = item(0)
                sum.Cells(r, 2).Value = nums(idx)
                sum.Cells(r, 3).Value = item(2)
                r += 1
                idx += 1
            Next
            sum.Range("A1:C1").Merge() : sum.Range("A1:C1").Font.Bold = True
            sum.Range("A1:C1").Font.Size = 16
            sum.Range("A9:C9").Font.Bold = True
            sum.Range("A9:C9").Interior.Color = RGB(215, 226, 242)
            sum.Range("B10:B22").NumberFormat = "0.000000"
            sum.Columns("A").ColumnWidth = 37 : sum.Columns("B").ColumnWidth = 39 : sum.Columns("C").ColumnWidth = 17
            If capturePath <> "" AndAlso File.Exists(capturePath) Then
                Dim shape As Object = sum.Shapes.AddPicture(capturePath, 0, -1, 540, 22, 460, 310)
                FreeCom(shape)
            End If

            Dim headings() As String = {"Occurrence path", "Part name", "Mass (kg)", "CG X (mm)", "CG Y (mm)", "CG Z (mm)", "Bodies", "Status"}
            For col As Integer = 0 To 7
                comps.Cells(1, col + 1).Value = headings(col)
            Next
            Dim nr As Integer = 2
            For Each cr As CompRec In SortedRecords()
                comps.Cells(nr, 1).Value = cr.Path : comps.Cells(nr, 2).Value = cr.PartName
                comps.Cells(nr, 3).Value = cr.Mass
                If cr.Mass > 0 Then
                    For i As Integer = 0 To 2
                        comps.Cells(nr, i + 4).Value = cr.CG(i)
                    Next
                End If
                comps.Cells(nr, 7).Value = cr.Cnt : comps.Cells(nr, 8).Value = cr.Status
                nr += 1
            Next
            comps.Range("A1:H1").Font.Bold = True
            comps.Range("A1:H1").Interior.Color = RGB(215, 226, 242)
            comps.Columns("A").ColumnWidth = 54 : comps.Columns("B").ColumnWidth = 28
            comps.Columns("C:F").ColumnWidth = 19 : comps.Columns("G:H").ColumnWidth = 18
            comps.Range("C2:F" & Math.Max(2, nr - 1).ToString()).NumberFormat = "0.000000"
            If nr > 2 Then comps.Range("A1:H" & (nr - 1).ToString()).AutoFilter()
            warns.Cells(1, 1).Value = "STATUS" : warns.Cells(1, 2).Value = If(isIncomplete, "INCOMPLETE", "MEASURED")
            warns.Cells(2, 1).Value = "Warnings / exclusions"
            Dim wr As Integer = 3
            For Each item As String In warnings
                warns.Cells(wr, 1).Value = item : wr += 1
            Next
            warns.Cells(wr + 1, 1).Value = "Not a collision/overlap correction. Verify active reference sets and WAVE duplication."
            warns.Columns("A").ColumnWidth = 105 : warns.Columns("B").ColumnWidth = 25
            warns.Range("A1:B2").Font.Bold = True
            book.SaveAs(fileName, 51)  ' Excel Open XML Workbook (*.xlsx)
        Finally
            If book IsNot Nothing Then
                Try : book.Close(False) : Catch : End Try
            End If
            If excel IsNot Nothing Then
                Try : excel.Quit() : Catch : End Try
            End If
            FreeCom(warns) : FreeCom(comps) : FreeCom(sum)
            FreeCom(book) : FreeCom(books) : FreeCom(excel)
            ' Some Office interop property chains produce temporary COM proxies.
            GC.Collect() : GC.WaitForPendingFinalizers()
        End Try
    End Sub

    Private Function SortedRecords() As List(Of CompRec)
        Dim items As New List(Of CompRec)(records.Values)
        items.Sort(Function(a, b) String.Compare(a.Path, b.Path, StringComparison.Ordinal))
        Return items
    End Function

    Private Sub ExportCsv(xlsxPath As String)
        Dim root As String = Path.Combine(Path.GetDirectoryName(xlsxPath), Path.GetFileNameWithoutExtension(xlsxPath))
        Using sw As New StreamWriter(root & "_Summary.csv", False, New UTF8Encoding(True))
            sw.WriteLine("Assembly," & Csv(top.Leaf))
            sw.WriteLine("Date," & Csv(DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")))
            sw.WriteLine("Status," & If(isIncomplete, "INCOMPLETE", "MEASURED"))
            sw.WriteLine("Metric,Value,Units")
            For Each item As String() In SummaryRows()
                sw.WriteLine(Csv(item(0)) & "," & Csv(item(1)) & "," & Csv(item(2)))
            Next
        End Using
        Using sw As New StreamWriter(root & "_Components.csv", False, New UTF8Encoding(True))
            sw.WriteLine("Occurrence Path,Part,Mass kg,CG X mm,CG Y mm,CG Z mm,Bodies,Status")
            For Each r As CompRec In SortedRecords()
                sw.WriteLine(Csv(r.Path) & "," & Csv(r.PartName) & "," & F(r.Mass) & "," &
                   F(r.CG(0)) & "," & F(r.CG(1)) & "," & F(r.CG(2)) & "," & r.Cnt & "," & Csv(r.Status))
            Next
        End Using
        Using sw As New StreamWriter(root & "_Warnings.csv", False, New UTF8Encoding(True))
            sw.WriteLine("Warning")
            For Each w As String In warnings : sw.WriteLine(Csv(w)) : Next
        End Using
    End Sub

    Private Sub CreateNXTable()
        If sess.Parts.Work Is Nothing OrElse sess.Parts.Work.Tag <> top.Tag Then
            Warn("Native table skipped: make the displayed TOP assembly the Work Part first.", False)
            Return
        End If
        If MessageBox.Show("Optional native NX tabular note is created in the current WORK part." & vbCrLf &
              "For a drawing: activate the desired drawing sheet first." & vbCrLf &
              "For modeling: select a location in the modeling view (availability depends on NX drafting context)." & vbCrLf &
              "No part will be saved. Continue to placement selection?", "Native NX table",
              MessageBoxButtons.OKCancel) <> DialogResult.OK Then Return
        Dim view As NXOpen.View = Nothing
        Dim loc As Point3d
        Dim response As Selection.DialogResponse = UI.GetUI().SelectionManager.SelectScreenPosition(
                         "Pick the insertion location for the mass-properties table", view, loc)
        If response <> Selection.DialogResponse.Pick Then
            Say("Native table cancelled by user.") : Return
        End If
        Dim sec As UFTabnot.SectionPrefs
        uf.Tabnot.AskDefaultSectionPrefs(sec)
        Dim orig() As Double = {loc.X, loc.Y, loc.Z}
        Dim nt As Tag = Tag.Null
        uf.Tabnot.Create(sec, orig, nt)
        Try
        ' Remove any template row/column created by default preferences.
        Dim n As Integer = 0
        uf.Tabnot.AskNmRows(nt, n)
        For k As Integer = 1 To n
            Dim row As Tag
            uf.Tabnot.AskNthRow(nt, 0, row)
            uf.Tabnot.RemoveRow(row) : uf.Obj.DeleteObject(row)
        Next
        uf.Tabnot.AskNmColumns(nt, n)
        For k As Integer = 1 To n
            Dim col As Tag
            uf.Tabnot.AskNthColumn(nt, 0, col)
            uf.Tabnot.RemoveColumn(col) : uf.Obj.DeleteObject(col)
        Next
        Dim cols(2) As Tag
        Dim partLengthPerMm As Double = If(top.PartUnits = Part.Units.Inches, 1.0 / 25.4, 1.0)
        For i As Integer = 0 To 2
            uf.Tabnot.CreateColumn(If(i = 0, 70.0, 38.0) * partLengthPerMm, cols(i))
            uf.Tabnot.AddColumn(nt, cols(i), UFConstants.UF_TABNOT_APPEND)
        Next
        Dim rows As New List(Of String())()
        rows.Add(New String() {"NX ASSEMBLY MASS / INERTIA", If(isIncomplete, "INCOMPLETE", "MEASURED"), ""})
        rows.Add(New String() {"Property", "Value", "Units"})
        rows.AddRange(SummaryRows())
        For Each vals As String() In rows
            Dim row As Tag
            uf.Tabnot.CreateRow(9.0 * partLengthPerMm, row)
            uf.Tabnot.AddRow(nt, row, UFConstants.UF_TABNOT_APPEND)
            For j As Integer = 0 To 2
                Dim cell As Tag
                uf.Tabnot.AskCellAtRowCol(row, cols(j), cell)
                uf.Tabnot.SetCellText(cell, vals(j))
            Next
        Next
        Say("Native tabular note created (UNSAVED). Tag: " & nt.ToString())
        Catch
            ' A failure after Create must not leave a partially built table behind.
            Try : uf.Obj.DeleteObject(nt) : Catch : End Try
            Throw
        End Try
    End Sub

    Private Function Good(x As Double) As Boolean
        Return Not Double.IsNaN(x) AndAlso Not Double.IsInfinity(x)
    End Function
    Private Function F(x As Double) As String
        Return x.ToString("0.000000", CultureInfo.InvariantCulture)
    End Function
    Private Function Csv(t As String) As String
        Return ChrW(34) & t.Replace(ChrW(34), ChrW(34) & ChrW(34)) & ChrW(34)
    End Function
    Private Sub Warn(msg As String, incomplete As Boolean)
        warnings.Add(msg)
        If incomplete Then isIncomplete = True
        Say("WARNING: " & msg)
    End Sub
    Private Sub Say(msg As String)
        lw.WriteLine(msg)
    End Sub
    Private Sub FreeCom(ByVal o As Object)
        If o Is Nothing Then Return
        Try
            If Marshal.IsComObject(o) Then Marshal.FinalReleaseComObject(o)
        Catch
        End Try
    End Sub
    Public Function GetUnloadOption(dummy As String) As Integer
        Return CInt(Session.LibraryUnloadOption.Immediately)
    End Function
End Module
