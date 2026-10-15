' NX 2306 - Assembly mass, CG and inertia report with image and NX table
' Version 1.0 - 2026-10-15
'
' HOW TO USE
' 1. Open and fully load the assembly in NX 2306. Make its top-level part
'    BOTH the displayed part and work part. Arrange the desired image view.
' 2. Tools > Journal > Play > select this .vb file.
' 3. Choose an existing output folder, scope and NX-table options.
' 4. Review the values, then choose Export Excel + image.
'
' Requires Windows NX and installed desktop Microsoft Excel for .xlsx output.
' No Excel interop DLL reference or third-party package is needed.
' The NX table is a real snapshot tabular-note annotation, with its own Undo
' mark. It is optional; the results grid and Listing Window always show data.
' No geometry, density, WCS, component placement or view orientation is edited.
' The part is never saved or closed. Repeated runs create separate reports
' and, when enabled, another table. NX Undo removes the new table only.
'
' OUTPUTS: Assembly_CG_MI.xlsx (embedded PNG), Assembly_View.png,
' Assembly_CG_MI.csv, Body_Inventory.csv and Run_Log.txt in a unique run folder.
' Mass = kg; CG = mm in assembly absolute coordinates; inertia = kg*mm^2.
' Both CG-centered inertia parallel to ABS axes and ABS-origin inertia are
' reported. Ixy/Ixz/Iyz are SIGNED TENSOR entries: Ixy=-integral(x*y dm),
' with coordinates relative to the stated origin. They are not unsigned
' products of inertia. Positive integral products are the negatives.
'
' SCOPE: solid body occurrences in the currently loaded assembly
' representation, including hidden bodies by default; suppression/reference
' sets apply. This is geometry-based, not a physical BOM or mass override.
' Fully load geometry. No automatic density fallback is used. Sheets/curves
' and bodies tagged NX_EXTERIOR_RESULT=v1 by the earlier exterior journal
' are excluded. Review untagged simplified/duplicate geometry yourself.
' The image is the current graphics view, not a rendered image of the scope.
' Hidden bodies included in measurement may not appear in the image.
'
' VALIDATION: reviewed against mirrored Siemens-authored API documentation
' (primarily NX2022.1). Not compiled or executed in NX 2306 here. Test a
' known-density block, repeated rotated components and a translated/rotated
' WCS against NX Measure before using this for engineering release.
' Pure numeric startup checks run when played; they do not verify NX APIs.
'
' FIRST LOCAL CHECKS
' - 100 mm cube from ABS (0,0,0) to (100,100,100), density 7850 kg/m^3:
'   mass=7.85 kg; CG=(50,50,50) mm; Ixx=Iyy=Izz at CG=13083.333333
'   kg*mm^2; ABS-origin Ixx=Iyy=Izz=52333.333333; origin Ixy=-19625.
' - Move/rotate WCS without moving the cube: reported ABS values must not change.
' - Repeat a component at a different placement: count both occurrences.
' - Unload an eligible component: stop rather than show an incomplete total.
' - Check NX table and Excel match the grid; PNG is embedded, not linked.
' - Undo table creation: geometry and WCS must remain unchanged.
'
' 1 kg*m^2 = 1,000,000 kg*mm^2. Excel and NX tables are snapshots, not live
' associative measures. Existing NX tables are not overwritten by reruns.
'
' APIs: UF.Modl.AskMassProps3d; UF.Csys.MapPoint; UF.Obj.CycleObjsInPart;
' Views.CreateImageExportBuilder; UF.Tabnot methods.
' References:
' https://nxopencsdocumentation.thescriptingengineer.com/NX2022_1/NXOpen.UF.UFModl.html
' https://www.ugapi.com/doc/ufun/uf_modl/uf_modl.html
' https://nxopencsdocumentation.thescriptingengineer.com/NX2022_1/NXOpen.UF.UFTabnot.html
' https://nxopencsdocumentation.thescriptingengineer.com/NX2022_1/NXOpen.Gateway.ImageExportBuilder.html
'
' Option Strict Off is required for late-bound desktop Excel COM only.
Option Strict Off
Option Explicit On

Imports System
Imports System.IO
Imports System.Text
Imports System.Collections.Generic
Imports System.Globalization
Imports System.Runtime.InteropServices
Imports NXOpen
Imports NXOpen.UF
Imports NXOpen.Assemblies
Imports NXOpen.Gateway
Imports NXOpen.Utilities
Imports Forms = System.Windows.Forms
Imports Drawing = System.Drawing

Module AssemblyCGMIExcel
    Private s As Session
    Private uf As UFSession
    Private part As Part
    Private messages As New List(Of String)()
    Private inventory As New List(Of InventoryItem)()
    Private folder As String = ""
    Private Const Accuracy As Double = 0.999

    Public Sub Main()
        s = Session.GetSession()
        uf = UFSession.GetUFSession()
        s.ListingWindow.Open()
        messages.Clear() : inventory.Clear() : folder = ""
        Try
            NumericSelfChecks()
            part = s.Parts.Work
            If part Is Nothing Then Throw New InvalidOperationException("Open the assembly first.")
            If s.Parts.Display Is Nothing OrElse part.Tag <> s.Parts.Display.Tag Then Throw New InvalidOperationException("Make the displayed top-level assembly the work part. In-place work components are not supported.")
            If Not part.IsFullyLoaded Then Throw New InvalidOperationException("Fully load the work part first.")
            Dim options As New OptionsForm()
            Using options
                If options.ShowDialog() <> Forms.DialogResult.OK Then Return
                If Not Directory.Exists(options.OutputFolder) Then Throw New DirectoryNotFoundException("Choose an existing output folder.")
                folder = Path.Combine(Path.GetFullPath(options.OutputFolder), "NX_CG_MI_" & DateTime.Now.ToString("yyyyMMdd_HHmmss") & "_" & Guid.NewGuid().ToString("N").Substring(0, 6))
                Directory.CreateDirectory(folder)
                Log("ASSEMBLY MASS / CG / INERTIA REPORT")
                Log("Time: " & DateTime.Now.ToString("s"))
                Log("Part: " & part.Leaf & " | " & part.FullPath)
                Log("NXOpen version: " & GetType(NXObject).Assembly.GetName().Version.ToString())
                Log("Scope: " & If(options.VisibleOnly, "visible solid occurrences in loaded representation", "all solid occurrences in loaded representation, including hidden"))
                Dim root As Component = part.ComponentAssembly.RootComponent
                If root IsNot Nothing Then
                    For Each child As Component In root.GetChildren()
                        CheckLoadedTree(child, options.VisibleOnly, False)
                    Next
                End If
                Dim bodies As List(Of Body) = CollectBodies(options.VisibleOnly, options.ExcludeExterior)
                WriteBodyInventory()
                If bodies.Count = 0 Then Throw New InvalidOperationException("No eligible solid body occurrences were found.")
                Dim result As ReportResult = MeasureAssembly(bodies)
                Dim reportRows As List(Of Metric) = MakeMetrics(result, bodies.Count)
                WriteCsv(reportRows)
                PrintTable(reportRows)
                If Not ReviewResults(reportRows, result, options.VisibleOnly) Then
                    Log("Excel export cancelled. CSV reports were kept; no NX table was created.")
                    Return
                End If

                ' Capture before creating the new NX table so it cannot cover the assembly.
                Dim imageFile As String = Path.Combine(folder, "Assembly_View.png")
                Dim imageReady As Boolean = False
                Try
                    CaptureImage(imageFile)
                    imageReady = True
                    Log("Assembly image captured from current graphics view.")
                Catch ex As Exception
                    Log("IMAGE FAILED: " & ex.Message)
                End Try
                If options.AddNXTable Then
                    Try
                        CreateNXTable(reportRows, options.TableOriginMm)
                        Log("NX tabular-note snapshot created. Work part is not saved.")
                    Catch ex As Exception
                        Log("NX TABLE FAILED: " & ex.Message)
                    End Try
                End If
                If imageReady Then
                    Try
                        ExportExcel(reportRows, result, imageFile, options.VisibleOnly)
                        Log("Excel report saved with embedded assembly image: " & Path.Combine(folder, "Assembly_CG_MI.xlsx"))
                        Forms.MessageBox.Show("Excel report saved with the assembly image." & vbCrLf & folder & vbCrLf & vbCrLf & "The NX part has not been saved. See Run_Log.txt for any table warnings.", "NX CG / MI report", Forms.MessageBoxButtons.OK, Forms.MessageBoxIcon.Information)
                    Catch ex As Exception
                        Log("EXCEL FAILED: " & ex.ToString())
                        Forms.MessageBox.Show("Excel export failed: " & ex.Message & vbCrLf & vbCrLf & "The measured values, CSV and PNG remain available in:" & vbCrLf & folder & vbCrLf & "Desktop Microsoft Excel must be installed.", "NX CG / MI report", Forms.MessageBoxButtons.OK, Forms.MessageBoxIcon.Warning)
                    End Try
                Else
                    Forms.MessageBox.Show("The image could not be captured, so the Excel-with-image export was not completed. Measurements and CSV are available in:" & vbCrLf & folder & vbCrLf & "See Run_Log.txt.", "NX CG / MI report", Forms.MessageBoxButtons.OK, Forms.MessageBoxIcon.Warning)
                End If
            End Using
        Catch ex As Exception
            Log("FAILED: " & ex.ToString())
            Forms.MessageBox.Show(ex.Message & vbCrLf & "See the NX Listing Window.", "NX CG / MI report", Forms.MessageBoxButtons.OK, Forms.MessageBoxIcon.Error)
        Finally
            If folder.Length > 0 Then
                Try
                    File.WriteAllText(Path.Combine(folder, "Run_Log.txt"), String.Join(vbCrLf, messages.ToArray()), New UTF8Encoding(False))
                Catch ex As Exception
                    s.ListingWindow.WriteLine("Cannot save run log: " & ex.Message)
                End Try
            End If
        End Try
    End Sub

    Private Sub CheckLoadedTree(comp As Component, visibleOnly As Boolean, ancestorHidden As Boolean)
        If comp.IsSuppressed OrElse comp.ReferenceSet = comp.EmptyPartRefsetName Then Return
        Dim hidden As Boolean = ancestorHidden OrElse comp.IsBlanked
        If visibleOnly AndAlso hidden Then Return
        Dim proto As Part = TryCast(comp.Prototype, Part)
        If proto Is Nothing OrElse Not proto.IsFullyLoaded Then Throw New InvalidOperationException("Fully load component geometry before reporting: " & ComponentPath(comp))
        For Each child As Component In comp.GetChildren()
            CheckLoadedTree(child, visibleOnly, hidden)
        Next
    End Sub

    Private Function CollectBodies(visibleOnly As Boolean, excludeExterior As Boolean) As List(Of Body)
        Dim result As New List(Of Body)()
        Dim seen As New HashSet(Of Tag)()
        Dim current As Tag = Tag.Null
        Dim errors As New List(Of String)()
        Do
            uf.Obj.CycleObjsInPart(part.Tag, UFConstants.UF_solid_type, current)
            If current = Tag.Null Then Exit Do
            Dim item As InventoryItem = Nothing
            Try
                Dim body As Body = TryCast(NXObjectManager.Get(current), Body)
                If body Is Nothing OrElse Not seen.Add(body.Tag) Then Continue Do
                item = New InventoryItem()
                item.TagText = body.Tag.ToString() : item.BodyName = body.Name
                item.OccurrencePath = If(body.OwningComponent Is Nothing, part.Leaf, ComponentPath(body.OwningComponent))
                inventory.Add(item)
                If Not body.IsSolidBody Then
                    item.Status = "EXCLUDED" : item.Reason = "not a solid body" : Continue Do
                End If
                If excludeExterior AndAlso IsEarlierExterior(body) Then
                    item.Status = "EXCLUDED" : item.Reason = "tagged exterior result NX_EXTERIOR_RESULT=v1" : Continue Do
                End If
                If Not EligibleOccurrence(body.OwningComponent, visibleOnly) Then
                    item.Status = "EXCLUDED" : item.Reason = "suppressed, empty reference set or hidden by selected scope" : Continue Do
                End If
                If visibleOnly AndAlso (body.IsBlanked OrElse part.Layers.GetState(body.Layer) = NXOpen.Layer.State.Hidden) Then
                    item.Status = "EXCLUDED" : item.Reason = "body blanked or layer hidden" : Continue Do
                End If
                If visibleOnly AndAlso body.IsOccurrence Then
                    Dim prototypeBody As Body = TryCast(body.Prototype, Body)
                    If prototypeBody IsNot Nothing AndAlso prototypeBody.OwningPart.Layers.GetState(prototypeBody.Layer) = NXOpen.Layer.State.Hidden Then
                        item.Status = "EXCLUDED" : item.Reason = "prototype body layer hidden" : Continue Do
                    End If
                End If
                Dim rho As Double = 0.0
                uf.Modl.AskBodyDensity(body.Tag, UFModl.DensityUnits.KilogramsMeters, rho)
                If Not Finite(rho) OrElse rho <= 0.0 Then Throw New InvalidOperationException("Missing/nonpositive body density at " & item.OccurrencePath & ", body tag " & item.TagText)
                item.DensityKgM3 = rho : item.Status = "INCLUDED" : item.Reason = "NX body density; occurrence tag retained"
                result.Add(body)
            Catch ex As Exception
                If item IsNot Nothing Then
                    item.Status = "ERROR" : item.Reason = ex.Message
                End If
                errors.Add("Body tag " & current.ToString() & ": " & ex.Message)
            End Try
        Loop
        If errors.Count > 0 Then
            For Each value As String In errors
                Log("BODY ERROR: " & value)
            Next
            WriteBodyInventory()
            Throw New InvalidOperationException("One or more solid bodies could not be validated. No partial assembly total is reported. See Body_Inventory.csv and the Listing Window.")
        End If
        Log("Included solid body occurrences: " & result.Count.ToString())
        Log("Densities are read from NX bodies. Verify assigned material/density before relying on the mass report; no default steel density is injected.")
        Return result
    End Function

    Private Function EligibleOccurrence(comp As Component, visibleOnly As Boolean) As Boolean
        Dim current As Component = comp
        While current IsNot Nothing
            If current.IsSuppressed OrElse current.ReferenceSet = current.EmptyPartRefsetName Then Return False
            If visibleOnly AndAlso current.IsBlanked Then Return False
            current = current.Parent
        End While
        Return True
    End Function

    Private Function ComponentPath(comp As Component) As String
        Dim names As New List(Of String)()
        Dim current As Component = comp
        While current IsNot Nothing
            names.Insert(0, If(String.IsNullOrWhiteSpace(current.Name), current.DisplayName, current.Name))
            current = current.Parent
        End While
        Return String.Join("/", names.ToArray())
    End Function

    Private Function IsEarlierExterior(body As Body) As Boolean
        If AttributeValue(body, "NX_EXTERIOR_RESULT") = "v1" Then Return True
        If body.IsOccurrence Then Return AttributeValue(TryCast(body.Prototype, NXObject), "NX_EXTERIOR_RESULT") = "v1"
        Return False
    End Function

    Private Function AttributeValue(obj As NXObject, title As String) As String
        If obj Is Nothing Then Return ""
        If obj.HasUserAttribute(title, NXObject.AttributeType.String, -1) Then Return obj.GetStringUserAttribute(title, -1)
        Return ""
    End Function

    Private Function MeasureAssembly(bodies As List(Of Body)) As ReportResult
        Dim tags As New List(Of Tag)()
        For Each body As Body In bodies
            tags.Add(body.Tag)
        Next
        Dim values(46) As Double, statistics(12) As Double, accuracyValues(10) As Double
        accuracyValues(0) = Accuracy
        ' Analysis 1 = solid; unit system 4 = kg and m. Density argument is
        ' ignored for solid analysis; each body's stored density is used.
        uf.Modl.AskMassProps3d(tags.ToArray(), tags.Count, 1, 4, 0.0, 1, accuracyValues, values, statistics)
        For Each index As Integer In New Integer() {1, 2, 3, 4, 5, 12, 13, 14, 31, 32, 33, 46}
            If Not Finite(values(index)) Then Throw New InvalidOperationException("NX returned a nonfinite mass-property value at index " & index.ToString())
        Next
        If values(2) <= 0.0 OrElse values(1) <= 0.0 Then Throw New InvalidOperationException("Assembly mass/volume is not positive. Check solid geometry and densities.")
        Dim result As New ReportResult()
        result.MassKg = values(2) : result.VolumeMm3 = values(1) * 1000000000.0
        result.DensityKgM3 = values(46)
        result.PrincipalKgMm2 = New Double() {values(31) * 1000000.0, values(32) * 1000000.0, values(33) * 1000000.0}

        ' Map WCS origin and unit bases to ABS without changing the user's WCS.
        ' MapPoint uses native part length units; UF mass results use meters.
        Dim wcsOriginNative() As Double = MapWcsToAbsolute(New Double() {0.0, 0.0, 0.0})
        Dim rotation(2, 2) As Double
        For j As Integer = 0 To 2
            Dim basis(2) As Double : basis(j) = 1.0
            Dim mapped() As Double = MapWcsToAbsolute(basis)
            For i As Integer = 0 To 2
                rotation(i, j) = mapped(i) - wcsOriginNative(i)
            Next
        Next
        CheckOrthonormal(rotation, "WCS-to-absolute rotation")
        Dim nativePerMeter As Double = If(part.PartUnits = BasePart.Units.Inches, 1000.0 / 25.4, 1000.0)
        Dim cgNative() As Double = MapWcsToAbsolute(New Double() {values(3) * nativePerMeter, values(4) * nativePerMeter, values(5) * nativePerMeter})
        Dim mmPerNative As Double = If(part.PartUnits = BasePart.Units.Inches, 25.4, 1.0)
        result.CgMm = New Double() {cgNative(0) * mmPerNative, cgNative(1) * mmPerNative, cgNative(2) * mmPerNative}

        ' Reconstruct the physical centroidal tensor from principal moments
        ' and principal axis directions. This avoids assuming a product-of-
        ' inertia array order/sign. Eigenvector signs do not affect the dyads.
        Dim axes(2, 2) As Double
        For k As Integer = 0 To 2
            For i As Integer = 0 To 2
                axes(i, k) = values(22 + 3 * k + i)
            Next
        Next
        CheckOrthonormal(axes, "NX principal axes")
        Dim tensorWcs(2, 2) As Double
        For k As Integer = 0 To 2
            If values(31 + k) < 0.0 Then Throw New InvalidOperationException("NX returned a negative principal moment.")
            For i As Integer = 0 To 2
                For j As Integer = 0 To 2
                    tensorWcs(i, j) += values(31 + k) * axes(i, k) * axes(j, k)
                Next
            Next
        Next
        Dim scale As Double = Math.Max(0.000000000001, Math.Abs(values(12)) + Math.Abs(values(13)) + Math.Abs(values(14)))
        For i As Integer = 0 To 2
            If Math.Abs(tensorWcs(i, i) - values(12 + i)) > scale * 0.0001 Then Throw New InvalidOperationException("Principal-axis reconstruction does not match NX centroidal moments. Verify the installed API output convention before use.")
        Next
        result.InertiaCg = RotateTensor(tensorWcs, rotation)
        For i As Integer = 0 To 2
            For j As Integer = 0 To 2
                result.InertiaCg(i, j) *= 1000000.0
            Next
        Next
        result.InertiaOrigin = ShiftFromCg(result.InertiaCg, result.MassKg, result.CgMm)
        Log("Coordinates mapped from the current WCS to assembly ABS; the WCS was not changed.")
        Log("Inertia tensor reconstructed from principal axes/moments and checked against NX centroidal diagonals.")
        Return result
    End Function

    Private Function MapWcsToAbsolute(value() As Double) As Double()
        Dim mapped(2) As Double
        uf.Csys.MapPoint(UFConstants.UF_CSYS_ROOT_WCS_COORDS, value, UFConstants.UF_CSYS_ROOT_COORDS, mapped)
        Return mapped
    End Function

    Private Function RotateTensor(tensor(,) As Double, rotation(,) As Double) As Double(,)
        Dim result(2, 2) As Double
        For i As Integer = 0 To 2
            For j As Integer = 0 To 2
                For a As Integer = 0 To 2
                    For b As Integer = 0 To 2
                        result(i, j) += rotation(i, a) * tensor(a, b) * rotation(j, b)
                    Next
                Next
            Next
        Next
        Return result
    End Function

    Private Function ShiftFromCg(tensor(,) As Double, mass As Double, cg() As Double) As Double(,)
        Dim result(2, 2) As Double
        Dim distanceSquared As Double = cg(0) * cg(0) + cg(1) * cg(1) + cg(2) * cg(2)
        For i As Integer = 0 To 2
            For j As Integer = 0 To 2
                result(i, j) = tensor(i, j) + mass * (If(i = j, distanceSquared, 0.0) - cg(i) * cg(j))
            Next
        Next
        Return result
    End Function

    Private Sub CheckOrthonormal(matrix(,) As Double, label As String)
        For a As Integer = 0 To 2
            For b As Integer = 0 To 2
                Dim dot As Double = 0.0
                For i As Integer = 0 To 2
                    If Not Finite(matrix(i, a)) Then Throw New InvalidOperationException(label & " contains invalid values.")
                    dot += matrix(i, a) * matrix(i, b)
                Next
                If Math.Abs(dot - If(a = b, 1.0, 0.0)) > 0.0001 Then Throw New InvalidOperationException(label & " is not orthonormal.")
            Next
        Next
    End Sub

    Private Function MakeMetrics(value As ReportResult, count As Integer) As List(Of Metric)
        Dim rows As New List(Of Metric)()
        rows.Add(New Metric("Mass", value.MassKg, "kg", "Selected solid occurrences"))
        Dim axisNames() As String = {"X", "Y", "Z"}
        For i As Integer = 0 To 2
            rows.Add(New Metric("CG " & axisNames(i), value.CgMm(i), "mm", "Assembly absolute coordinates"))
        Next
        AddTensorMetrics(rows, value.InertiaCg, "CG", "Axes through assembly CG, parallel to ABS")
        AddTensorMetrics(rows, value.InertiaOrigin, "ABS origin", "Assembly absolute origin and axes")
        For i As Integer = 0 To 2
            rows.Add(New Metric("Principal I" & (i + 1).ToString(), value.PrincipalKgMm2(i), "kg*mm^2", "At assembly CG; NX principal order"))
        Next
        rows.Add(New Metric("Volume", value.VolumeMm3, "mm^3", "Sum of selected solid volumes"))
        rows.Add(New Metric("Solid occurrences", CDbl(count), "count", "Repeated instances counted separately"))
        rows.Add(New Metric("Combined density", value.DensityKgM3, "kg/m^3", "NX effective density"))
        Return rows
    End Function

    Private Sub AddTensorMetrics(rows As List(Of Metric), tensor(,) As Double, location As String, reference As String)
        rows.Add(New Metric("Ixx @ " & location, tensor(0, 0), "kg*mm^2", reference))
        rows.Add(New Metric("Iyy @ " & location, tensor(1, 1), "kg*mm^2", reference))
        rows.Add(New Metric("Izz @ " & location, tensor(2, 2), "kg*mm^2", reference))
        rows.Add(New Metric("Ixy @ " & location, tensor(0, 1), "kg*mm^2", reference & "; signed tensor entry"))
        rows.Add(New Metric("Ixz @ " & location, tensor(0, 2), "kg*mm^2", reference & "; signed tensor entry"))
        rows.Add(New Metric("Iyz @ " & location, tensor(1, 2), "kg*mm^2", reference & "; signed tensor entry"))
    End Sub

    Private Sub PrintTable(rows As List(Of Metric))
        Log("Quantity".PadRight(26) & "Value".PadLeft(22) & "  Unit")
        For Each row As Metric In rows
            Log(row.Name.PadRight(26) & row.Value.ToString("G12", CultureInfo.InvariantCulture).PadLeft(22) & "  " & row.UnitText)
        Next
    End Sub

    Private Function ReviewResults(rows As List(Of Metric), result As ReportResult, visibleOnly As Boolean) As Boolean
        Using form As New Forms.Form()
            form.Text = "Assembly mass / CG / MI - " & part.Leaf
            form.Size = New Drawing.Size(1050, 710)
            form.StartPosition = Forms.FormStartPosition.CenterScreen
            Dim grid As New Forms.DataGridView()
            grid.Dock = Forms.DockStyle.Fill : grid.ReadOnly = True : grid.AllowUserToAddRows = False
            grid.RowHeadersVisible = False : grid.AutoSizeColumnsMode = Forms.DataGridViewAutoSizeColumnsMode.Fill
            grid.Columns.Add("Quantity", "Quantity") : grid.Columns.Add("Value", "Value")
            grid.Columns.Add("Unit", "Unit") : grid.Columns.Add("Reference", "Reference")
            grid.Columns(0).FillWeight = 24 : grid.Columns(1).FillWeight = 22
            grid.Columns(2).FillWeight = 12 : grid.Columns(3).FillWeight = 65
            For Each row As Metric In rows
                grid.Rows.Add(row.Name, row.Value.ToString("G12", CultureInfo.InvariantCulture), row.UnitText, row.Reference)
            Next
            Dim note As New Forms.Label()
            note.Dock = Forms.DockStyle.Top : note.Height = 68 : note.Padding = New Forms.Padding(12, 8, 12, 0)
            note.Text = "Scope: " & If(visibleOnly, "visible", "all represented, including hidden") & " solid occurrences. Image uses the current graphics view." & vbCrLf & "Ixy/Ixz/Iyz are signed tensor entries. Verify body densities. Export creates a snapshot report."
            Dim buttons As New Forms.FlowLayoutPanel()
            buttons.Dock = Forms.DockStyle.Bottom : buttons.Height = 50 : buttons.FlowDirection = Forms.FlowDirection.RightToLeft
            Dim export As New Forms.Button()
            export.Text = "Export Excel + image" : export.AutoSize = True : export.Height = 34
            export.DialogResult = Forms.DialogResult.OK
            Dim cancel As New Forms.Button()
            cancel.Text = "Close without Excel" : cancel.AutoSize = True : cancel.Height = 34
            cancel.DialogResult = Forms.DialogResult.Cancel
            buttons.Controls.Add(export) : buttons.Controls.Add(cancel)
            form.Controls.Add(grid) : form.Controls.Add(note) : form.Controls.Add(buttons)
            form.AcceptButton = export : form.CancelButton = cancel
            Return form.ShowDialog() = Forms.DialogResult.OK
        End Using
    End Function

    Private Sub CaptureImage(filename As String)
        Dim builder As ImageExportBuilder = Nothing
        Try
            builder = part.Views.CreateImageExportBuilder()
            builder.FileFormat = ImageExportBuilder.FileFormats.Png
            builder.FileName = filename
            builder.RegionMode = False
            builder.BackgroundOption = ImageExportBuilder.BackgroundOptions.Original
            builder.Commit()
        Finally
            If builder IsNot Nothing Then builder.Destroy()
        End Try
        If Not File.Exists(filename) AndAlso File.Exists(filename & ".png") Then File.Move(filename & ".png", filename)
        If Not File.Exists(filename) OrElse New FileInfo(filename).Length = 0 Then Throw New IOException("NX did not create a nonempty PNG.")
        Using image As Drawing.Image = Drawing.Image.FromFile(filename)
            If image.Width < 10 OrElse image.Height < 10 Then Throw New IOException("The captured image has invalid dimensions.")
        End Using
    End Sub

    Private Sub CreateNXTable(metrics As List(Of Metric), originMm() As Double)
        If part.IsReadOnly OrElse Not part.HasWriteAccess Then Throw New InvalidOperationException("The work part is read-only. Excel reporting can still proceed.")
        Dim mark As Session.UndoMarkId = s.SetUndoMark(Session.MarkVisibility.Visible, "Assembly CG MI report table")
        Try
            Dim prefs As New UFTabnot.SectionPrefs()
            uf.Tabnot.AskDefaultSectionPrefs(prefs)
            Dim origin() As Double = {ToNative(originMm(0)), ToNative(originMm(1)), ToNative(originMm(2))}
            Dim table As Tag = Tag.Null
            uf.Tabnot.Create(prefs, origin, table)
            Dim nCols As Integer = 0, nRows As Integer = 0
            uf.Tabnot.AskNmColumns(table, nCols) : uf.Tabnot.AskNmRows(table, nRows)
            While nCols > 4
                Dim col As Tag = Tag.Null
                uf.Tabnot.AskNthColumn(table, nCols - 1, col)
                uf.Tabnot.RemoveColumn(col) : uf.Obj.DeleteObject(col) : nCols -= 1
            End While
            While nCols < 4
                Dim col As Tag = Tag.Null
                uf.Tabnot.CreateColumn(ToNative(40.0), col)
                uf.Tabnot.AddColumn(table, col, UFConstants.UF_TABNOT_APPEND) : nCols += 1
            End While
            Dim wantedRows As Integer = metrics.Count + 2
            While nRows > wantedRows
                Dim row As Tag = Tag.Null
                uf.Tabnot.AskNthRow(table, nRows - 1, row)
                uf.Tabnot.RemoveRow(row) : uf.Obj.DeleteObject(row) : nRows -= 1
            End While
            While nRows < wantedRows
                Dim row As Tag = Tag.Null
                uf.Tabnot.CreateRow(ToNative(8.0), row)
                uf.Tabnot.AddRow(table, row, UFConstants.UF_TABNOT_APPEND) : nRows += 1
            End While
            Dim widths() As Double = {58.0, 45.0, 25.0, 105.0}
            Dim columns(3) As Tag
            For c As Integer = 0 To 3
                uf.Tabnot.AskNthColumn(table, c, columns(c))
                uf.Tabnot.SetColumnWidth(columns(c), ToNative(widths(c)))
            Next
            For r As Integer = 0 To wantedRows - 1
                Dim row As Tag = Tag.Null
                uf.Tabnot.AskNthRow(table, r, row)
                uf.Tabnot.SetRowHeight(row, ToNative(8.0))
                Dim text() As String
                If r = 0 Then
                    text = New String() {"MASS / CG / MI - " & part.Leaf, "", "", ""}
                ElseIf r = 1 Then
                    text = New String() {"Quantity", "Value", "Unit", "Reference"}
                Else
                    Dim metric As Metric = metrics(r - 2)
                    text = New String() {metric.Name, metric.Value.ToString("G10", CultureInfo.InvariantCulture), metric.UnitText, metric.Reference}
                End If
                For c As Integer = 0 To 3
                    Dim cell As Tag = Tag.Null
                    uf.Tabnot.AskCellAtRowCol(row, columns(c), cell)
                    Dim cellPrefs As New UFTabnot.CellPrefs()
                    uf.Tabnot.AskCellPrefs(cell, cellPrefs)
                    cellPrefs.text_height = ToNative(3.0) : cellPrefs.is_a_formula = False
                    uf.Tabnot.SetCellPrefs(cell, cellPrefs)
                    uf.Tabnot.SetCellText(cell, text(c))
                Next
            Next
            Dim titleRow As Tag = Tag.Null, titleFirst As Tag = Tag.Null, titleLast As Tag = Tag.Null
            uf.Tabnot.AskNthRow(table, 0, titleRow)
            uf.Tabnot.AskCellAtRowCol(titleRow, columns(0), titleFirst)
            uf.Tabnot.AskCellAtRowCol(titleRow, columns(3), titleLast)
            uf.Tabnot.MergeCells(titleFirst, titleLast)
            uf.Tabnot.Update(table)
            Dim obj As NXObject = TryCast(NXObjectManager.Get(table), NXObject)
            If obj IsNot Nothing Then obj.SetName("CG_MI_" & DateTime.Now.ToString("yyyyMMdd_HHmmss"))
        Catch
            s.UndoToMark(mark, Nothing)
            Throw
        End Try
    End Sub

    Private Sub ExportExcel(metrics As List(Of Metric), result As ReportResult, imageFile As String, visibleOnly As Boolean)
        Dim app As Object = Nothing, workbook As Object = Nothing
        Dim objects As New ComObjects()
        Dim saved As Boolean = False
        Try
            app = objects.Keep(CreateObject("Excel.Application"))
            app.Visible = False : app.DisplayAlerts = False
            Dim books As Object = objects.Keep(app.Workbooks)
            workbook = objects.Keep(books.Add(-4167)) ' xlWBATWorksheet: one worksheet
            Dim sheets As Object = objects.Keep(workbook.Worksheets)
            Dim sheet As Object = objects.Keep(sheets.Item(1))
            sheet.Name = "Assembly report"
            Dim allCells As Object = objects.Keep(sheet.Cells)
            Dim allFont As Object = objects.Keep(allCells.Font)
            allFont.Name = "Calibri" : allFont.Size = 11
            Dim title As Object = objects.Keep(sheet.Range("A1:L2"))
            title.Merge() : title.Value2 = "ASSEMBLY MASS / CG / INERTIA"
            Dim titleFont As Object = objects.Keep(title.Font)
            titleFont.Size = 20 : titleFont.Bold = True : titleFont.Color = Drawing.ColorTranslator.ToOle(Drawing.Color.White)
            Dim titleFill As Object = objects.Keep(title.Interior)
            titleFill.Color = Drawing.ColorTranslator.ToOle(Drawing.Color.FromArgb(26, 60, 78))
            Dim partLabel As Object = objects.Keep(sheet.Range("A3:L3"))
            partLabel.Merge() : partLabel.NumberFormat = "@"
            partLabel.Value2 = part.Leaf & "  |  " & DateTime.Now.ToString("yyyy-MM-dd HH:mm")
            Dim headers(0, 3) As Object
            headers(0, 0) = "Quantity" : headers(0, 1) = "Value" : headers(0, 2) = "Unit" : headers(0, 3) = "Reference"
            Dim headerRange As Object = objects.Keep(sheet.Range("A5:D5"))
            headerRange.Value2 = headers
            Dim headerFont As Object = objects.Keep(headerRange.Font)
            headerFont.Bold = True : headerFont.Color = Drawing.ColorTranslator.ToOle(Drawing.Color.White)
            Dim headerFill As Object = objects.Keep(headerRange.Interior)
            headerFill.Color = Drawing.ColorTranslator.ToOle(Drawing.Color.FromArgb(31, 117, 124))
            Dim data(metrics.Count - 1, 3) As Object
            For i As Integer = 0 To metrics.Count - 1
                data(i, 0) = metrics(i).Name : data(i, 1) = metrics(i).Value
                data(i, 2) = metrics(i).UnitText : data(i, 3) = metrics(i).Reference
            Next
            Dim finalRow As Integer = 5 + metrics.Count
            Dim dataRange As Object = objects.Keep(sheet.Range("A6:D" & finalRow.ToString()))
            dataRange.Value2 = data : dataRange.RowHeight = 24
            Dim values As Object = objects.Keep(sheet.Range("B6:B" & finalRow.ToString()))
            values.NumberFormat = "0.000000"
            For i As Integer = 0 To metrics.Count - 1
                Dim magnitude As Double = Math.Abs(metrics(i).Value)
                If magnitude >= 1000000000.0 OrElse (magnitude > 0.0 AndAlso magnitude < 0.000001) Then
                    Dim numericCell As Object = objects.Keep(sheet.Range("B" & (i + 6).ToString()))
                    numericCell.NumberFormat = "0.000000E+00"
                End If
                If metrics(i).UnitText = "count" Then
                    Dim countCell As Object = objects.Keep(sheet.Range("B" & (i + 6).ToString()))
                    countCell.NumberFormat = "0"
                End If
            Next
            Dim references As Object = objects.Keep(sheet.Range("D6:D" & finalRow.ToString()))
            references.WrapText = True
            Dim colA As Object = objects.Keep(sheet.Range("A:A")) : colA.ColumnWidth = 24
            Dim colB As Object = objects.Keep(sheet.Range("B:B")) : colB.ColumnWidth = 21
            Dim colC As Object = objects.Keep(sheet.Range("C:C")) : colC.ColumnWidth = 13
            Dim colD As Object = objects.Keep(sheet.Range("D:D")) : colD.ColumnWidth = 40
            Dim photoCols As Object = objects.Keep(sheet.Range("E:L")) : photoCols.ColumnWidth = 10
            Dim caption As Object = objects.Keep(sheet.Range("F5:L5"))
            caption.Merge() : caption.Value2 = "ASSEMBLY IMAGE - CURRENT NX VIEW"
            Dim captionFont As Object = objects.Keep(caption.Font) : captionFont.Bold = True
            Dim anchor As Object = objects.Keep(sheet.Range("F7"))
            Dim shapes As Object = objects.Keep(sheet.Shapes)
            Dim picture As Object = objects.Keep(shapes.AddPicture(imageFile, 0, -1, CDbl(anchor.Left), CDbl(anchor.Top), -1, -1))
            picture.LockAspectRatio = -1
            picture.Width = 450.0
            If CDbl(picture.Height) > 330.0 Then picture.Height = 330.0

            Dim notes() As String = {
                "CG: assembly absolute coordinates in mm. CG inertia: axes through CG, parallel to absolute X/Y/Z. Absolute-origin inertia is listed separately.",
                "Ixy/Ixz/Iyz are signed physical tensor entries, e.g. Ixy = -integral((x-cx)*(y-cy) dm) at CG. Positive inertia products have the opposite sign.",
                "Scope: " & If(visibleOnly, "visible", "all, including hidden") & " solid occurrences in the loaded representation. Suppression/reference sets apply. This is not a BOM mass-override report.",
                "The PNG captures the current graphics view. It can omit hidden measured bodies and can show excluded visible objects. No camera/WCS/visibility was changed.",
                "Uses stored NX body densities without fallback. Review density assignments. Untagged simplified/duplicate geometry must be excluded by the user.",
                "Snapshot only. Geometry changes do not update this workbook or the NX table. Re-run the journal to generate a new report."
            }
            For i As Integer = 0 To notes.Length - 1
                Dim rowNumber As Integer = finalRow + 3 + i
                Dim noteRange As Object = objects.Keep(sheet.Range("A" & rowNumber.ToString() & ":L" & rowNumber.ToString()))
                noteRange.Merge() : noteRange.Value2 = notes(i) : noteRange.WrapText = True : noteRange.RowHeight = 32
                Dim font As Object = objects.Keep(noteRange.Font) : font.Size = 10
            Next
            Dim page As Object = objects.Keep(sheet.PageSetup)
            page.Orientation = 2 : page.Zoom = False : page.FitToPagesWide = 1 : page.FitToPagesTall = False
            page.PrintArea = "$A$1:$L$" & (finalRow + 8).ToString()

            Dim audit As Object = objects.Keep(sheets.Add(Type.Missing, sheet))
            audit.Name = "Body inventory"
            Dim auditTitle As Object = objects.Keep(audit.Range("A1:F1"))
            auditTitle.Merge() : auditTitle.Value2 = "BODY OCCURRENCES AND DENSITY USED"
            Dim auditTitleFont As Object = objects.Keep(auditTitle.Font) : auditTitleFont.Size = 15 : auditTitleFont.Bold = True
            Dim auditHeaders(0, 5) As Object
            auditHeaders(0, 0) = "Occurrence path" : auditHeaders(0, 1) = "Body name" : auditHeaders(0, 2) = "NX tag"
            auditHeaders(0, 3) = "Density kg/m^3" : auditHeaders(0, 4) = "Status" : auditHeaders(0, 5) = "Reason"
            Dim auditHeaderRange As Object = objects.Keep(audit.Range("A3:F3"))
            auditHeaderRange.Value2 = auditHeaders
            Dim auditFont As Object = objects.Keep(auditHeaderRange.Font) : auditFont.Bold = True
            If inventory.Count > 0 Then
                Dim auditValues(inventory.Count - 1, 5) As Object
                For i As Integer = 0 To inventory.Count - 1
                    Dim item As InventoryItem = inventory(i)
                    auditValues(i, 0) = SafeExcelText(item.OccurrencePath) : auditValues(i, 1) = SafeExcelText(item.BodyName)
                    auditValues(i, 2) = item.TagText
                    auditValues(i, 3) = If(item.Status = "INCLUDED", CType(item.DensityKgM3, Object), CType(Nothing, Object))
                    auditValues(i, 4) = item.Status : auditValues(i, 5) = SafeExcelText(item.Reason)
                Next
                Dim auditRange As Object = objects.Keep(audit.Range("A4:F" & (inventory.Count + 3).ToString()))
                auditRange.Value2 = auditValues
            End If
            Dim auditColumns As Object = objects.Keep(audit.Range("A:F")) : auditColumns.ColumnWidth = 25
            Dim paths As Object = objects.Keep(audit.Range("A:A")) : paths.ColumnWidth = 60
            Dim reasons As Object = objects.Keep(audit.Range("F:F")) : reasons.ColumnWidth = 70
            Dim densityRange As Object = objects.Keep(audit.Range("D:D")) : densityRange.NumberFormat = "0.000000"
            sheet.Activate()
            workbook.SaveAs(Path.Combine(folder, "Assembly_CG_MI.xlsx"), 51)
            saved = True
        Finally
            If workbook IsNot Nothing Then
                Try
                    workbook.Close(False)
                Catch
                End Try
            End If
            If app IsNot Nothing Then
                Try
                    app.Quit()
                Catch
                End Try
            End If
            objects.ReleaseAll()
        End Try
        If Not saved OrElse Not File.Exists(Path.Combine(folder, "Assembly_CG_MI.xlsx")) Then Throw New IOException("Excel did not save the workbook.")
    End Sub

    Private Function SafeExcelText(value As String) As String
        If value Is Nothing Then Return ""
        If value.Length > 0 AndAlso "=+-@".IndexOf(value(0)) >= 0 Then Return "'" & value
        Return value
    End Function

    Private Sub WriteCsv(rows As List(Of Metric))
        Dim text As New StringBuilder("Quantity,Value,Unit,Reference" & vbCrLf)
        For Each row As Metric In rows
            text.AppendLine(Csv(row.Name) & "," & row.Value.ToString("R", CultureInfo.InvariantCulture) & "," & Csv(row.UnitText) & "," & Csv(row.Reference))
        Next
        File.WriteAllText(Path.Combine(folder, "Assembly_CG_MI.csv"), text.ToString(), New UTF8Encoding(True))
    End Sub

    Private Sub WriteBodyInventory()
        If folder.Length = 0 Then Return
        Dim text As New StringBuilder("Occurrence path,Body name,NX tag,Density kg/m^3,Status,Reason" & vbCrLf)
        For Each item As InventoryItem In inventory
            text.AppendLine(Csv(item.OccurrencePath) & "," & Csv(item.BodyName) & "," & Csv(item.TagText) & "," & If(item.Status = "INCLUDED", item.DensityKgM3.ToString("R", CultureInfo.InvariantCulture), "") & "," & Csv(item.Status) & "," & Csv(item.Reason))
        Next
        File.WriteAllText(Path.Combine(folder, "Body_Inventory.csv"), text.ToString(), New UTF8Encoding(True))
    End Sub

    Private Function Csv(value As String) As String
        Return ChrW(34) & If(value, "").Replace(ChrW(34), ChrW(34) & ChrW(34)) & ChrW(34)
    End Function
    Private Function ToNative(mm As Double) As Double
        Return If(part.PartUnits = BasePart.Units.Inches, mm / 25.4, mm)
    End Function
    Private Function Finite(value As Double) As Boolean
        Return Not Double.IsNaN(value) AndAlso Not Double.IsInfinity(value)
    End Function
    Private Sub Log(value As String)
        messages.Add(value) : s.ListingWindow.WriteLine(value)
    End Sub

    Private Sub NumericSelfChecks()
        Dim tensor(,) As Double = {{2.0, 0.0, 0.0}, {0.0, 3.0, 0.0}, {0.0, 0.0, 4.0}}
        Dim rotation(,) As Double = {{0.0, -1.0, 0.0}, {1.0, 0.0, 0.0}, {0.0, 0.0, 1.0}}
        CheckOrthonormal(rotation, "startup rotation")
        Dim rotated As Double(,) = RotateTensor(tensor, rotation)
        If Math.Abs(rotated(0, 0) - 3.0) > 0.0000001 OrElse Math.Abs(rotated(1, 1) - 2.0) > 0.0000001 Then Throw New InvalidOperationException("Tensor rotation self-check failed.")
        Dim c As Double = Math.Sqrt(0.5)
        Dim diagonalRotation(,) As Double = {{c, -c, 0.0}, {c, c, 0.0}, {0.0, 0.0, 1.0}}
        Dim diagonalTensor As Double(,) = RotateTensor(tensor, diagonalRotation)
        If Math.Abs(diagonalTensor(0, 1) + 0.5) > 0.0000001 Then Throw New InvalidOperationException("Off-diagonal rotation/sign self-check failed.")
        Dim shifted As Double(,) = ShiftFromCg(tensor, 2.0, New Double() {3.0, 4.0, 0.0})
        If Math.Abs(shifted(0, 0) - 34.0) > 0.0000001 OrElse Math.Abs(shifted(1, 1) - 21.0) > 0.0000001 OrElse Math.Abs(shifted(2, 2) - 54.0) > 0.0000001 OrElse Math.Abs(shifted(0, 1) + 24.0) > 0.0000001 Then Throw New InvalidOperationException("Parallel-axis/sign self-check failed.")
    End Sub

    Public Function GetUnloadOption(dummy As String) As Integer
        Return CInt(Session.LibraryUnloadOption.Immediately)
    End Function

    Private Class Metric
        Public Name As String, UnitText As String, Reference As String
        Public Value As Double
        Public Sub New(nameValue As String, number As Double, units As String, referenceValue As String)
            Name = nameValue : Value = number : UnitText = units : Reference = referenceValue
        End Sub
    End Class
    Private Class ReportResult
        Public MassKg As Double, VolumeMm3 As Double, DensityKgM3 As Double
        Public CgMm() As Double, PrincipalKgMm2() As Double
        Public InertiaCg(,) As Double, InertiaOrigin(,) As Double
    End Class
    Private Class InventoryItem
        Public OccurrencePath As String = "", BodyName As String = "", TagText As String = ""
        Public Status As String = "ERROR", Reason As String = "inspection did not complete"
        Public DensityKgM3 As Double
    End Class
    Private Class ComObjects
        Private values As New List(Of Object)()
        Public Function Keep(value As Object) As Object
            For Each existing As Object In values
                If Object.ReferenceEquals(existing, value) Then Return value
            Next
            values.Add(value) : Return value
        End Function
        Public Sub ReleaseAll()
            For i As Integer = values.Count - 1 To 0 Step -1
                Try
                    If Marshal.IsComObject(values(i)) Then Marshal.FinalReleaseComObject(values(i))
                Catch
                End Try
            Next
            values.Clear()
        End Sub
    End Class
    Private Class OptionsForm
        Inherits Forms.Form
        Private output As New Forms.TextBox()
        Private visible As New Forms.CheckBox()
        Private exterior As New Forms.CheckBox()
        Private table As New Forms.CheckBox()
        Private coordinates As New List(Of Forms.NumericUpDown)()
        Public ReadOnly Property OutputFolder As String
            Get
                Return output.Text.Trim()
            End Get
        End Property
        Public ReadOnly Property VisibleOnly As Boolean
            Get
                Return visible.Checked
            End Get
        End Property
        Public ReadOnly Property ExcludeExterior As Boolean
            Get
                Return exterior.Checked
            End Get
        End Property
        Public ReadOnly Property AddNXTable As Boolean
            Get
                Return table.Checked
            End Get
        End Property
        Public ReadOnly Property TableOriginMm As Double()
            Get
                Return New Double() {CDbl(coordinates(0).Value), CDbl(coordinates(1).Value), CDbl(coordinates(2).Value)}
            End Get
        End Property
        Public Sub New()
            Text = "NX 2306 - Assembly CG / MI report"
            ClientSize = New Drawing.Size(700, 360)
            StartPosition = Forms.FormStartPosition.CenterScreen
            FormBorderStyle = Forms.FormBorderStyle.FixedDialog : MaximizeBox = False
            Dim instruction As New Forms.Label()
            instruction.Text = "Arrange the assembly view before running. Excel embeds the current graphics image."
            instruction.SetBounds(16, 16, 665, 30) : Controls.Add(instruction)
            Dim label As New Forms.Label()
            label.Text = "Existing output folder" : label.SetBounds(16, 55, 600, 22) : Controls.Add(label)
            output.Text = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments)
            output.SetBounds(16, 80, 550, 26)
            Dim browse As New Forms.Button()
            browse.Text = "Browse..." : browse.SetBounds(580, 78, 100, 29)
            AddHandler browse.Click, AddressOf BrowseFolder
            visible.Text = "Visible solid occurrences only (unchecked includes hidden bodies)"
            visible.SetBounds(16, 125, 665, 26) : visible.Checked = False
            exterior.Text = "Exclude tagged exterior results from the earlier exterior-export journal"
            exterior.SetBounds(16, 155, 665, 26) : exterior.Checked = True
            table.Text = "Create an NX tabular-note snapshot (modifies annotations only; no save)"
            table.SetBounds(16, 185, 665, 26) : table.Checked = True
            Dim origin As New Forms.Label()
            origin.Text = "NX table origin: absolute X / Y / Z in mm"
            origin.SetBounds(16, 225, 665, 22) : Controls.Add(origin)
            For i As Integer = 0 To 2
                Dim number As New Forms.NumericUpDown()
                number.DecimalPlaces = 3 : number.Minimum = -1000000D : number.Maximum = 1000000D
                number.SetBounds(16 + 170 * i, 252, 155, 26)
                coordinates.Add(number) : Controls.Add(number)
            Next
            Dim accept As New Forms.Button()
            accept.Text = "Calculate and review" : accept.SetBounds(380, 310, 180, 32)
            accept.DialogResult = Forms.DialogResult.OK
            Dim cancel As New Forms.Button()
            cancel.Text = "Cancel" : cancel.SetBounds(575, 310, 105, 32)
            cancel.DialogResult = Forms.DialogResult.Cancel
            Controls.AddRange(New Forms.Control() {output, browse, visible, exterior, table, accept, cancel})
            AcceptButton = accept : CancelButton = cancel
        End Sub
        Private Sub BrowseFolder(sender As Object, e As EventArgs)
            Using dialog As New Forms.FolderBrowserDialog()
                dialog.SelectedPath = output.Text
                If dialog.ShowDialog() = Forms.DialogResult.OK Then output.Text = dialog.SelectedPath
            End Using
        End Sub
    End Class
End Module
