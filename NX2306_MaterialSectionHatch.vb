' NX 2306 - Material-based assembly-section hatch pilot
' Run as Tools > Journal > Play (VB.NET)
' IMPORTANT: Review on a COPY of the drawing first.
' This journal edits ONLY hatches whose annotation associations resolve to
' exactly one assembly component occurrence. It never guesses from 2D position.
' Default is PREVIEW ONLY. Change ApplyChanges to True after validating report.
'
Option Strict Off
Imports System
Imports System.IO
Imports System.Text
Imports System.Collections.Generic
Imports NXOpen
Imports NXOpen.Annotations
Imports NXOpen.Assemblies
Imports NXOpen.Drawings

Module NX2306_MaterialSectionHatch

    Private Const ApplyChanges As Boolean = False
    Private Const SectionViewsOnly As Boolean = True
    Private Const PreserveManualHatches As Boolean = True

    Private Class HatchRule
        Public Family As String
        Public Angle As Double
        Public Spacing As Double
        Public Pattern As String
        Public Sub New(f As String, a As Double, d As Double, Optional p As String = "")
            Family = f : Angle = a : Spacing = d : Pattern = p
        End Sub
    End Class

    Public Sub Main(ByVal args() As String)
        Dim session As Session = Session.GetSession()
        Dim log As ListingWindow = session.ListingWindow
        log.Open()
        Dim part As Part = session.Parts.Work
        If part Is Nothing Then
            log.WriteLine("Open the assembly drawing as the Work Part.")
            Return
        End If

        Dim output As String = Path.Combine(Path.GetTempPath(), "NX_SectionHatch_" & DateTime.Now.ToString("yyyyMMdd_HHmmss") & ".csv")
        Dim rows As New List(Of String)
        rows.Add("HatchTag,View,OccurrencePath,PartNumber,Material,Rule,AngleDeg,Spacing,Status,Details")
        Dim done As Integer = 0, skipped As Integer = 0, changed As Integer = 0
        Dim mark As Session.UndoMarkId = session.SetUndoMark(Session.MarkVisibility.Visible, "Material section hatching")
        log.WriteLine("Material Hatch Pilot (NX 2306) - " & If(ApplyChanges, "APPLY", "PREVIEW") & " mode")

        Try
            For Each h As Hatch In part.Annotations.Hatches.ToArray()
                Dim viewName As String = ""
                Dim allowed As Boolean = Not SectionViewsOnly
                Try
                    For Each v As NXOpen.View In h.GetViews()
                        If v Is Nothing Then Continue For
                        viewName = v.Name
                        If TypeOf v Is SectionView Then allowed = True
                    Next
                Catch
                End Try
                If Not allowed Then Continue For

                Dim occurrence As Component = Nothing
                Dim status As String = ""
                Dim detail As String = ""
                Dim pathName As String = "", pn As String = "", material As String = ""
                Dim rule As HatchRule = Nothing

                Dim owners As New Dictionary(Of String, Component)(StringComparer.OrdinalIgnoreCase)
                Try
                    For i As Integer = 1 To h.NumberOfAssociativities
                        Dim a As Associativity = h.GetAssociativity(i)
                        If a Is Nothing Then Continue For
                        AddOwner(a.FirstObject, owners)
                        AddOwner(a.SecondObject, owners)
                    Next
                Catch ex As Exception
                    detail = "Association read issue: " & ex.Message
                End Try

                If owners.Count = 1 Then
                    For Each pair In owners
                        occurrence = pair.Value
                    Next
                ElseIf owners.Count > 1 Then
                    status = "AMBIGUOUS_OWNER"
                Else
                    status = "NO_DIRECT_OWNER"
                End If

                If occurrence IsNot Nothing Then
                    pathName = GetPath(occurrence)
                    Dim prototype As Part = TryCast(occurrence.Prototype, Part)
                    If prototype Is Nothing Then
                        status = "PROTOTYPE_NOT_LOADED"
                    Else
                        pn = ReadFirst(prototype, New String() {"DB_PART_NO", "PART_NUMBER", "PART_NO"})
                        If pn = "" Then pn = prototype.Leaf
                        material = ReadFirst(prototype, New String() {"MBD_MATERIAL_CLASS", "MATERIAL_ID", "MATERIAL", "MATL"})
                        Dim overrideMat As String = ReadFirst(occurrence, New String() {"MBD_MATERIAL_OVERRIDE"})
                        If overrideMat <> "" Then material = overrideMat
                        If PreserveManualHatches AndAlso ReadFirst(h, New String() {"MBD_HATCH_LOCK"}).ToUpperInvariant() = "TRUE" Then
                            status = "LOCKED"
                        ElseIf material = "" Then
                            status = "MATERIAL_MISSING"
                        Else
                            rule = ChooseRule(material)
                            If rule Is Nothing Then
                                status = "RULE_NOT_FOUND"
                            Else
                                status = "READY"
                                If ApplyChanges Then
                                    Try
                                        ApplyStyle(part, h, rule)
                                        status = "APPLIED"
                                        changed += 1
                                    Catch ex As Exception
                                        status = "APPLY_FAILED"
                                        detail = ex.Message
                                    End Try
                                End If
                            End If
                        End If
                    End If
                End If
                done += 1
                If status <> "READY" AndAlso status <> "APPLIED" Then skipped += 1
                Dim angle As String = If(rule Is Nothing, "", rule.Angle.ToString(Globalization.CultureInfo.InvariantCulture))
                Dim spacing As String = If(rule Is Nothing, "", rule.Spacing.ToString(Globalization.CultureInfo.InvariantCulture))
                rows.Add(String.Join(",", New String() {Csv(h.Tag.ToString()), Csv(viewName), Csv(pathName), Csv(pn), Csv(material), Csv(If(rule Is Nothing, "", rule.Family)), Csv(angle), Csv(spacing), Csv(status), Csv(detail)}))
                log.WriteLine(h.Tag.ToString() & " | " & viewName & " | " & pn & " | " & material & " | " & status)
            Next
            If ApplyChanges Then session.UpdateManager.DoUpdate(mark)
        Catch ex As Exception
            log.WriteLine("GENERAL ERROR: " & ex.ToString())
            If ApplyChanges Then
                Try
                    session.UndoToMark(mark, Nothing)
                    log.WriteLine("Undid this journal's changes.")
                Catch
                End Try
            End If
        Finally
            Try
                File.WriteAllLines(output, rows.ToArray(), Encoding.UTF8)
                log.WriteLine("CSV report: " & output)
            Catch ex As Exception
                log.WriteLine("CSV write failed: " & ex.Message)
            End Try
            log.WriteLine("Hatches examined: " & done & " | not ready/applied: " & skipped & " | changed: " & changed)
            log.WriteLine("NOTE: NO_DIRECT_OWNER requires a section-topology mapper before safe automatic editing.")
        End Try
    End Sub

    Private Sub AddOwner(ByVal obj As NXObject, ByVal owners As Dictionary(Of String, Component))
        If obj Is Nothing Then Return
        Dim comp As Component = TryCast(obj, Component)
        If comp Is Nothing Then comp = obj.OwningComponent
        If comp Is Nothing Then
            Dim proto As NXObject = obj.Prototype
            If proto IsNot Nothing Then comp = proto.OwningComponent
        End If
        If comp Is Nothing Then Return
        Dim key As String = GetPath(comp)
        If Not owners.ContainsKey(key) Then owners.Add(key, comp)
    End Sub

    Private Function GetPath(ByVal comp As Component) As String
        Dim names As New Stack(Of String)
        Dim current As Component = comp
        While current IsNot Nothing
            names.Push(current.DisplayName)
            current = current.Parent
        End While
        Return String.Join("/", names.ToArray())
    End Function

    Private Function ReadFirst(ByVal obj As NXObject, ByVal names As String()) As String
        For Each n As String In names
            Try
                If obj.HasUserAttribute(n, NXObject.AttributeType.String, -1) Then
                    Dim s As String = obj.GetStringUserAttribute(n, -1)
                    If Not String.IsNullOrWhiteSpace(s) Then Return s.Trim()
                End If
            Catch
            End Try
        Next
        Return ""
    End Function

    Private Function ChooseRule(ByVal rawMaterial As String) As HatchRule
        Dim s As String = rawMaterial.ToUpperInvariant().Replace("-", " ").Replace("_", " ").Trim()
        ' EXAMPLE presentation rules only; replace with company crosshatch standard.
        If s.Contains("ALUMIN") OrElse s.StartsWith("AL ") OrElse s.StartsWith("ALLOY AL") Then Return New HatchRule("ALUMINUM", 135.0, 2.5)
        If s.Contains("BRONZE") OrElse s.Contains("BRASS") OrElse s.Contains("COPPER") Then Return New HatchRule("COPPER_ALLOY", 45.0, 1.5)
        If s.Contains("CAST IRON") OrElse s.Contains("SG IRON") OrElse s.Contains("DUCTILE IRON") Then Return New HatchRule("CAST_IRON", 45.0, 2.0)
        If s.Contains("STEEL") OrElse s.Contains("AISI") OrElse s.Contains("EN8") OrElse s.Contains("EN19") Then Return New HatchRule("STEEL", 45.0, 2.5)
        If s.Contains("RUBBER") OrElse s.Contains("NBR") OrElse s.Contains("EPDM") Then Return New HatchRule("ELASTOMER", 135.0, 1.25)
        Return Nothing
    End Function

    Private Sub ApplyStyle(ByVal part As Part, ByVal h As Hatch, ByVal rule As HatchRule)
        ' Edit the existing native hatch only; do not build a new boundary.
        Dim builder As HatchBuilder = part.Annotations.Hatches.CreateHatchBuilder(h)
        Try
            ' The exact site style/pattern names are intentionally not assumed.
            Dim settings As Object = builder.HatchFillSettings
            settings.Angle = rule.Angle
            settings.Distance = rule.Spacing
            settings.ApplyToAllFromSameComp = False
            If rule.Pattern <> "" Then settings.Pattern = rule.Pattern
            builder.Commit()
        Finally
            builder.Destroy()
        End Try
    End Sub

    Private Function Csv(ByVal value As String) As String
        If value Is Nothing Then value = ""
        Return """" & value.Replace("""", """""") & """"
    End Function

    Public Function GetUnloadOption(ByVal dummy As String) As Integer
        Return CInt(Session.LibraryUnloadOption.Immediately)
    End Function
End Module
