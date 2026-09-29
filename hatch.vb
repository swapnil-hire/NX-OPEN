'==============================================================================
' NX 2306 - MATERIAL BASED SECTION HATCHING
' VB.NET / NXOpen Journal
'
' PURPOSE:
'   Automatically changes existing drawing section/cross-hatch patterns
'   according to the material of the associated component/body.
'
' MATERIAL SOURCE PRIORITY:
'   1. NX Physical Material assigned to associated body
'   2. Component instance material attribute
'   3. Component / prototype part material attribute
'   4. Associated object's owning part material attribute
'
' DEFAULT NX XHATCH PATTERNS:
'
'   Steel                         -> STEEL
'   Cast Iron                     -> IRON/GENERAL USE
'   Brass / Copper / Bronze       -> BRASS/COPPER
'   Aluminium / Magnesium         -> ALUMINUM/MAGNESIUM
'   Rubber / Plastic              -> RUBBER/PLASTIC
'   Refractory                    -> REFRACTORY
'   Marble / Slate / Glass        -> MARBLE/SLATE/GLASS
'   Lead / Zinc / Tin             -> LEAD
'
' IMPORTANT:
'   Material attribute aliases can be modified in MATERIAL_ATTRIBUTE_NAMES.
'
'   This journal modifies EXISTING Hatch objects.
'   It does not create new section hatches.
'
'==============================================================================

Option Strict On

Imports System
Imports System.IO
Imports System.Collections.Generic

Imports NXOpen
Imports NXOpen.Annotations
Imports NXOpen.Assemblies


Module NX2306_Material_Based_Hatching


    '==========================================================================
    ' NX SESSION
    '==========================================================================

    Private theSession As Session
    Private workPart As Part
    Private lw As ListingWindow


    '==========================================================================
    ' SETTINGS
    '==========================================================================

    Private Const DEFAULT_ANGLE As Double = 45.0

    Private Const DEFAULT_DISTANCE As Double = 3.0

    Private Const DEFAULT_SCALE As Double = 1.0


    'If True:
    'Unknown material = hatch is NOT changed.
    '
    'If False:
    'Unknown material = IRON/GENERAL USE pattern.
    Private Const SKIP_UNKNOWN_MATERIAL As Boolean = True


    '==========================================================================
    ' MATERIAL ATTRIBUTE NAMES
    '
    'Add your Teamcenter/NX attribute names here.
    '==========================================================================

    Private ReadOnly MATERIAL_ATTRIBUTE_NAMES As String() = {

        "MATERIAL",
        "DB_MATERIAL",
        "MATERIAL_SPEC",
        "MATERIAL SPEC",
        "MATERIAL_SPECIFICATION",
        "MATERIAL SPECIFICATION",
        "MATERIAL_GRADE",
        "MATERIAL GRADE",
        "MATL",
        "MATL_SPEC",
        "MATL SPEC",
        "GRADE"

    }


    '==========================================================================
    ' HATCH RULE CLASS
    '==========================================================================

    Private Class HatchRule

        Public Category As String

        Public Pattern As String

        Public Angle As Double

        Public Distance As Double

        Public Scale As Double


        Public Sub New(
            ByVal categoryValue As String,
            ByVal patternValue As String,
            ByVal angleValue As Double,
            ByVal distanceValue As Double,
            ByVal scaleValue As Double)

            Category = categoryValue

            Pattern = patternValue

            Angle = angleValue

            Distance = distanceValue

            Scale = scaleValue

        End Sub

    End Class


    '==========================================================================
    ' MATERIAL HATCH RULES
    '==========================================================================

    Private ReadOnly HatchRules _
        As New Dictionary(Of String, HatchRule)(
        StringComparer.OrdinalIgnoreCase) From {

        {
            "STEEL",
            New HatchRule(
                "STEEL",
                "STEEL",
                45.0,
                3.0,
                1.0)
        },

        {
            "CAST IRON",
            New HatchRule(
                "CAST IRON",
                "IRON/GENERAL USE",
                45.0,
                3.0,
                1.0)
        },

        {
            "IRON",
            New HatchRule(
                "IRON",
                "IRON/GENERAL USE",
                45.0,
                3.0,
                1.0)
        },

        {
            "BRASS/COPPER",
            New HatchRule(
                "BRASS/COPPER",
                "BRASS/COPPER",
                45.0,
                3.0,
                1.0)
        },

        {
            "ALUMINUM/MAGNESIUM",
            New HatchRule(
                "ALUMINUM/MAGNESIUM",
                "ALUMINUM/MAGNESIUM",
                45.0,
                3.0,
                1.0)
        },

        {
            "RUBBER/PLASTIC",
            New HatchRule(
                "RUBBER/PLASTIC",
                "RUBBER/PLASTIC",
                45.0,
                3.0,
                1.0)
        },

        {
            "REFRACTORY",
            New HatchRule(
                "REFRACTORY",
                "REFRACTORY",
                45.0,
                3.0,
                1.0)
        },

        {
            "MARBLE/SLATE/GLASS",
            New HatchRule(
                "MARBLE/SLATE/GLASS",
                "MARBLE/SLATE/GLASS",
                45.0,
                3.0,
                1.0)
        },

        {
            "LEAD",
            New HatchRule(
                "LEAD",
                "LEAD",
                45.0,
                3.0,
                1.0)
        }

    }


    '==========================================================================
    ' COUNTERS
    '==========================================================================

    Private totalHatches As Integer = 0

    Private changedHatches As Integer = 0

    Private skippedHatches As Integer = 0

    Private errorHatches As Integer = 0


    '==========================================================================
    ' MAIN
    '==========================================================================

    Public Sub Main()

        theSession = Session.GetSession()

        workPart = theSession.Parts.Work

        lw = theSession.ListingWindow

        lw.Open()


        lw.WriteLine("")

        lw.WriteLine(
            "============================================================")

        lw.WriteLine(
            "        NX MATERIAL BASED SECTION HATCHING")

        lw.WriteLine(
            "============================================================")

        lw.WriteLine("")


        If workPart Is Nothing Then

            lw.WriteLine(
                "ERROR: No work part is open.")

            Return

        End If


        Dim undoMark As Session.UndoMarkId =
            theSession.SetUndoMark(
                Session.MarkVisibility.Visible,
                "Material Based Section Hatching")


        Try

            ProcessAllHatches()


            Try

                theSession.UpdateManager.DoUpdate(
                    undoMark)

            Catch ex As Exception

                lw.WriteLine(
                    "Update warning: " &
                    ex.Message)

            End Try


        Catch ex As Exception

            lw.WriteLine("")

            lw.WriteLine(
                "FATAL ERROR:")

            lw.WriteLine(
                ex.Message)

        End Try


        PrintSummary()

    End Sub


    '==========================================================================
    ' PROCESS ALL HATCHES
    '==========================================================================

    Private Sub ProcessAllHatches()

        Dim foundAnyHatch As Boolean = False


        For Each hatch As Hatch _
            In workPart.Annotations.Hatches

            foundAnyHatch = True

            totalHatches += 1


            Try

                ProcessSingleHatch(
                    hatch,
                    totalHatches)

            Catch ex As Exception

                errorHatches += 1

                lw.WriteLine("")

                lw.WriteLine(
                    "[ERROR] Hatch " &
                    totalHatches.ToString())

                lw.WriteLine(
                    "Tag      : " &
                    hatch.Tag.ToString())

                lw.WriteLine(
                    "Message  : " &
                    ex.Message)

            End Try

        Next


        If Not foundAnyHatch Then

            lw.WriteLine(
                "No drawing hatch objects were found.")

        End If

    End Sub


    '==========================================================================
    ' PROCESS ONE HATCH
    '==========================================================================

    Private Sub ProcessSingleHatch(
        ByVal hatch As Hatch,
        ByVal hatchNumber As Integer)


        lw.WriteLine("")

        lw.WriteLine(
            "------------------------------------------------------------")

        lw.WriteLine(
            "HATCH " &
            hatchNumber.ToString())

        lw.WriteLine(
            "Tag      : " &
            hatch.Tag.ToString())


        '----------------------------------------------------------------------
        ' Find material
        '----------------------------------------------------------------------

        Dim materialSource As String = ""

        Dim materialName As String =
            GetMaterialFromHatch(
                hatch,
                materialSource)


        If String.IsNullOrWhiteSpace(
            materialName) Then


            lw.WriteLine(
                "Material : NOT FOUND")


            If SKIP_UNKNOWN_MATERIAL Then

                lw.WriteLine(
                    "Status   : SKIPPED")

                skippedHatches += 1

                Return

            Else

                materialName =
                    "IRON"

            End If

        End If


        lw.WriteLine(
            "Material : " &
            materialName)

        lw.WriteLine(
            "Source   : " &
            materialSource)


        '----------------------------------------------------------------------
        ' Convert actual material to generic hatch category
        '----------------------------------------------------------------------

        Dim category As String =
            NormalizeMaterial(
                materialName)


        lw.WriteLine(
            "Category : " &
            category)


        If String.IsNullOrWhiteSpace(
            category) Then

            lw.WriteLine(
                "Status   : MATERIAL NOT RECOGNIZED")

            skippedHatches += 1

            Return

        End If


        '----------------------------------------------------------------------
        ' Get hatch rule
        '----------------------------------------------------------------------

        Dim rule As HatchRule = Nothing


        If Not HatchRules.TryGetValue(
            category,
            rule) Then

            lw.WriteLine(
                "Status   : NO HATCH RULE")

            skippedHatches += 1

            Return

        End If


        '----------------------------------------------------------------------
        ' Change hatch
        '----------------------------------------------------------------------

        ApplyHatchRule(
            hatch,
            rule)


        changedHatches += 1


        lw.WriteLine(
            "Pattern  : " &
            rule.Pattern)

        lw.WriteLine(
            "Angle    : " &
            rule.Angle.ToString())

        lw.WriteLine(
            "Distance : " &
            rule.Distance.ToString())

        lw.WriteLine(
            "Scale    : " &
            rule.Scale.ToString())

        lw.WriteLine(
            "Status   : UPDATED")

    End Sub


    '==========================================================================
    ' GET MATERIAL FROM HATCH ASSOCIATIVITY
    '==========================================================================

    Private Function GetMaterialFromHatch(
        ByVal hatch As Hatch,
        ByRef sourceDescription As String) As String


        sourceDescription = ""


        Dim numberOfAssociativities As Integer = 0


        Try

            numberOfAssociativities =
                hatch.NumberOfAssociativities

        Catch

            Return ""

        End Try


        For i As Integer = 0 _
            To numberOfAssociativities - 1


            Dim assoc As Associativity = Nothing


            Try

                assoc =
                    hatch.GetAssociativity(i)


                If assoc Is Nothing Then
                    Continue For
                End If


                '--------------------------------------------------------------
                ' First associated object
                '--------------------------------------------------------------

                If assoc.FirstObject IsNot Nothing Then

                    Dim material As String =
                        GetMaterialFromNXObject(
                            assoc.FirstObject,
                            sourceDescription,
                            0)


                    If Not String.IsNullOrWhiteSpace(
                        material) Then

                        Return material

                    End If

                End If


                '--------------------------------------------------------------
                ' Second associated object
                '--------------------------------------------------------------

                If assoc.SecondObject IsNot Nothing Then

                    Dim material As String =
                        GetMaterialFromNXObject(
                            assoc.SecondObject,
                            sourceDescription,
                            0)


                    If Not String.IsNullOrWhiteSpace(
                        material) Then

                        Return material

                    End If

                End If


            Catch


            Finally

                If assoc IsNot Nothing Then

                    Try
                        assoc.Dispose()
                    Catch
                    End Try

                End If

            End Try

        Next


        Return ""

    End Function


    '==========================================================================
    ' GET MATERIAL FROM ASSOCIATED NX OBJECT
    '==========================================================================

    Private Function GetMaterialFromNXObject(
        ByVal obj As NXObject,
        ByRef sourceDescription As String,
        ByVal recursionLevel As Integer) As String


        If obj Is Nothing Then
            Return ""
        End If


        If recursionLevel > 5 Then
            Return ""
        End If


        '----------------------------------------------------------------------
        ' BODY / FACE / EDGE
        '----------------------------------------------------------------------

        Dim associatedBody As Body =
            GetBodyFromObject(
                obj)


        If associatedBody IsNot Nothing Then


            '------------------------------------------------------------------
            ' Physical Material
            '------------------------------------------------------------------

            Dim physicalMaterialName As String =
                GetPhysicalMaterialName(
                    associatedBody)


            If Not String.IsNullOrWhiteSpace(
                physicalMaterialName) Then

                sourceDescription =
                    "NX Physical Material"

                Return physicalMaterialName

            End If


            '------------------------------------------------------------------
            ' Body attribute
            '------------------------------------------------------------------

            Dim bodyMaterial As String =
                GetMaterialAttribute(
                    associatedBody)


            If Not String.IsNullOrWhiteSpace(
                bodyMaterial) Then

                sourceDescription =
                    "Body Attribute"

                Return bodyMaterial

            End If

        End If


        '----------------------------------------------------------------------
        ' COMPONENT
        '----------------------------------------------------------------------

        Dim component As Component = Nothing


        Try

            component =
                obj.OwningComponent

        Catch

            component = Nothing

        End Try


        If component IsNot Nothing Then


            '------------------------------------------------------------------
            ' Component instance attribute
            '------------------------------------------------------------------

            Dim instanceMaterial As String =
                GetMaterialFromComponentInstance(
                    component)


            If Not String.IsNullOrWhiteSpace(
                instanceMaterial) Then

                sourceDescription =
                    "Component Instance Attribute"

                Return instanceMaterial

            End If


            '------------------------------------------------------------------
            ' Component regular attribute
            '------------------------------------------------------------------

            Dim componentMaterial As String =
                GetMaterialAttribute(
                    component)


            If Not String.IsNullOrWhiteSpace(
                componentMaterial) Then

                sourceDescription =
                    "Component Attribute"

                Return componentMaterial

            End If


            '------------------------------------------------------------------
            ' Component prototype
            '------------------------------------------------------------------

            Try

                Dim prototypePart As BasePart =
                    TryCast(
                        component.Prototype,
                        BasePart)


                If prototypePart IsNot Nothing Then

                    Dim partMaterial As String =
                        GetMaterialAttribute(
                            prototypePart)


                    If Not String.IsNullOrWhiteSpace(
                        partMaterial) Then

                        sourceDescription =
                            "Component Prototype Part Attribute"

                        Return partMaterial

                    End If

                End If

            Catch

            End Try

        End If


        '----------------------------------------------------------------------
        ' OBJECT ATTRIBUTE
        '----------------------------------------------------------------------

        Dim objectMaterial As String =
            GetMaterialAttribute(
                obj)


        If Not String.IsNullOrWhiteSpace(
            objectMaterial) Then

            sourceDescription =
                "Associated Object Attribute"

            Return objectMaterial

        End If


        '----------------------------------------------------------------------
        ' PROTOTYPE OBJECT
        '
        ' Important for assembly occurrences and drafting geometry.
        '----------------------------------------------------------------------

        Try

            If obj.IsOccurrence Then

                Dim prototypeObject As NXObject =
                    TryCast(
                        obj.Prototype,
                        NXObject)


                If prototypeObject IsNot Nothing Then

                    Dim material As String =
                        GetMaterialFromNXObject(
                            prototypeObject,
                            sourceDescription,
                            recursionLevel + 1)


                    If Not String.IsNullOrWhiteSpace(
                        material) Then

                        Return material

                    End If

                End If

            End If

        Catch

        End Try


        '----------------------------------------------------------------------
        ' OWNING PART ATTRIBUTE
        '----------------------------------------------------------------------

        Try

            Dim ownerPart As BasePart =
                obj.OwningPart


            If ownerPart IsNot Nothing AndAlso
               ownerPart IsNot workPart Then


                Dim ownerMaterial As String =
                    GetMaterialAttribute(
                        ownerPart)


                If Not String.IsNullOrWhiteSpace(
                    ownerMaterial) Then

                    sourceDescription =
                        "Owning Part Attribute"

                    Return ownerMaterial

                End If

            End If

        Catch

        End Try


        Return ""

    End Function


    '==========================================================================
    ' GET BODY FROM OBJECT
    '==========================================================================

    Private Function GetBodyFromObject(
        ByVal obj As NXObject) As Body


        If obj Is Nothing Then
            Return Nothing
        End If


        '----------------------------------------------------------------------
        ' BODY
        '----------------------------------------------------------------------

        If TypeOf obj Is Body Then

            Return CType(
                obj,
                Body)

        End If


        '----------------------------------------------------------------------
        ' FACE
        '----------------------------------------------------------------------

        If TypeOf obj Is Face Then

            Try

                Return CType(
                    obj,
                    Face).GetBody()

            Catch

            End Try

        End If


        '----------------------------------------------------------------------
        ' EDGE
        '----------------------------------------------------------------------

        If TypeOf obj Is Edge Then

            Try

                Return CType(
                    obj,
                    Edge).GetBody()

            Catch

            End Try

        End If


        '----------------------------------------------------------------------
        ' OCCURRENCE PROTOTYPE
        '----------------------------------------------------------------------

        Try

            If obj.IsOccurrence Then

                Dim prototypeObj As NXObject =
                    TryCast(
                        obj.Prototype,
                        NXObject)


                If prototypeObj IsNot Nothing Then

                    If TypeOf prototypeObj Is Body Then

                        Return CType(
                            prototypeObj,
                            Body)

                    End If


                    If TypeOf prototypeObj Is Face Then

                        Return CType(
                            prototypeObj,
                            Face).GetBody()

                    End If


                    If TypeOf prototypeObj Is Edge Then

                        Return CType(
                            prototypeObj,
                            Edge).GetBody()

                    End If

                End If

            End If

        Catch

        End Try


        Return Nothing

    End Function


    '==========================================================================
    ' NX PHYSICAL MATERIAL
    '==========================================================================

    Private Function GetPhysicalMaterialName(
        ByVal inputBody As Body) As String


        If inputBody Is Nothing Then
            Return ""
        End If


        Try

            Dim targetBody As Body =
                inputBody


            'If this is an assembly occurrence,
            'query the prototype body.
            If inputBody.IsOccurrence Then

                Dim prototypeBody As Body =
                    TryCast(
                        inputBody.Prototype,
                        Body)


                If prototypeBody IsNot Nothing Then

                    targetBody =
                        prototypeBody

                End If

            End If


            Dim ownerPart As BasePart =
                targetBody.OwningPart


            If ownerPart Is Nothing Then
                Return ""
            End If


            Dim physicalMaterial As PhysicalMaterial =
                ownerPart.MaterialManager.PhysicalMaterials.
                AskMaterialOfObject(
                    targetBody)


            If physicalMaterial Is Nothing Then
                Return ""
            End If


            '------------------------------------------------------------------
            ' Material name
            '------------------------------------------------------------------

            Dim materialName As String = ""


            Try

                materialName =
                    physicalMaterial.Name

            Catch

            End Try


            If Not String.IsNullOrWhiteSpace(
                materialName) Then

                Return materialName

            End If


            '------------------------------------------------------------------
            ' Fallback to JournalIdentifier
            '
            ' Often:
            ' PhysicalMaterial[SAE 8620]
            '------------------------------------------------------------------

            Try

                Dim journalId As String =
                    physicalMaterial.JournalIdentifier


                If Not String.IsNullOrWhiteSpace(
                    journalId) Then


                    Dim startIndex As Integer =
                        journalId.IndexOf(
                            "["c)


                    Dim endIndex As Integer =
                        journalId.LastIndexOf(
                            "]"c)


                    If startIndex >= 0 AndAlso
                       endIndex > startIndex Then

                        Return journalId.Substring(
                            startIndex + 1,
                            endIndex -
                            startIndex -
                            1)

                    End If


                    Return journalId

                End If

            Catch

            End Try


        Catch

        End Try


        Return ""

    End Function


    '==========================================================================
    ' COMPONENT INSTANCE MATERIAL
    '==========================================================================

    Private Function GetMaterialFromComponentInstance(
        ByVal component As Component) As String


        If component Is Nothing Then
            Return ""
        End If


        Try

            Dim attributes() As NXObject.AttributeInformation =
                component.GetInstanceUserAttributes()


            For Each attribute As NXObject.AttributeInformation _
                In attributes


                If attribute.Unset Then
                    Continue For
                End If


                If IsMaterialAttributeName(
                    attribute.Title) Then


                    Dim value As String =
                        GetAttributeValue(
                            attribute)


                    If Not String.IsNullOrWhiteSpace(
                        value) Then

                        Return value.Trim()

                    End If

                End If

            Next


        Catch

        End Try


        Return ""

    End Function


    '==========================================================================
    ' READ MATERIAL ATTRIBUTE FROM NX OBJECT
    '==========================================================================

    Private Function GetMaterialAttribute(
        ByVal obj As NXObject) As String


        If obj Is Nothing Then
            Return ""
        End If


        Try

            Dim attributes() As NXObject.AttributeInformation =
                obj.GetUserAttributes()


            For Each attribute As NXObject.AttributeInformation _
                In attributes


                If attribute.Unset Then
                    Continue For
                End If


                If IsMaterialAttributeName(
                    attribute.Title) Then


                    Dim value As String =
                        GetAttributeValue(
                            attribute)


                    If Not String.IsNullOrWhiteSpace(
                        value) Then

                        Return value.Trim()

                    End If

                End If

            Next


        Catch

        End Try


        Return ""

    End Function


    '==========================================================================
    ' CHECK ATTRIBUTE NAME
    '==========================================================================

    Private Function IsMaterialAttributeName(
        ByVal attributeTitle As String) As Boolean


        If String.IsNullOrWhiteSpace(
            attributeTitle) Then

            Return False

        End If


        For Each materialAttribute As String _
            In MATERIAL_ATTRIBUTE_NAMES


            If String.Equals(
                attributeTitle.Trim(),
                materialAttribute.Trim(),
                StringComparison.OrdinalIgnoreCase) Then

                Return True

            End If

        Next


        Return False

    End Function


    '==========================================================================
    ' GET ATTRIBUTE VALUE
    '==========================================================================

    Private Function GetAttributeValue(
        ByVal attribute As NXObject.AttributeInformation) As String


        Try

            Select Case attribute.Type


                Case NXObject.AttributeType.String

                    Return attribute.StringValue


                Case NXObject.AttributeType.Integer

                    Return attribute.IntegerValue.ToString()


                Case NXObject.AttributeType.Real

                    Return attribute.RealValue.ToString()


                Case Else

                    If Not String.IsNullOrWhiteSpace(
                        attribute.StringValue) Then

                        Return attribute.StringValue

                    End If

            End Select


        Catch

        End Try


        Return ""

    End Function


    '==========================================================================
    ' NORMALIZE MATERIAL
    '==========================================================================

    Private Function NormalizeMaterial(
        ByVal rawMaterial As String) As String


        If String.IsNullOrWhiteSpace(
            rawMaterial) Then

            Return ""

        End If


        Dim material As String =
            rawMaterial.Trim().
            ToUpperInvariant()


        material =
            material.Replace("_", " ")

        material =
            material.Replace("-", " ")

        material =
            material.Replace("/", " ")


        '======================================================================
        ' CAST IRON
        ' Check before generic IRON.
        '======================================================================

        If material.Contains(
            "CAST IRON") OrElse

           material.Contains(
            "DUCTILE IRON") OrElse

           material.Contains(
            "NODULAR IRON") OrElse

           material.Contains(
            "SG IRON") OrElse

           material.Contains(
            "GREY IRON") OrElse

           material.Contains(
            "GRAY IRON") OrElse

           material.Contains(
            "GJL") OrElse

           material.Contains(
            "GJS") OrElse

           material.Contains(
            "GG25") OrElse

           material.Contains(
            "GG 25") OrElse

           material.Contains(
            "GGG") Then

            Return "CAST IRON"

        End If


        '======================================================================
        ' ALUMINIUM / ALUMINUM / MAGNESIUM
        '======================================================================

        If material.Contains(
            "ALUMIN") OrElse

           material.Contains(
            "ALSI") OrElse

           material.Contains(
            "AL SI") OrElse

           material.Contains(
            "MAGNESIUM") OrElse

           material.Contains(
            "MG ALLOY") Then

            Return "ALUMINUM/MAGNESIUM"

        End If


        '======================================================================
        ' COPPER / BRASS / BRONZE
        '======================================================================

        If material.Contains(
            "COPPER") OrElse

           material.Contains(
            "BRASS") OrElse

           material.Contains(
            "BRONZE") OrElse

           material.Contains(
            "GUNMETAL") OrElse

           material.Contains(
            "GUN METAL") Then

            Return "BRASS/COPPER"

        End If


        '======================================================================
        ' RUBBER / PLASTIC / POLYMER
        '======================================================================

        If material.Contains(
            "RUBBER") OrElse

           material.Contains(
            "PLASTIC") OrElse

           material.Contains(
            "POLYMER") OrElse

           material.Contains(
            "NYLON") OrElse

           material.Contains(
            "PTFE") OrElse

           material.Contains(
            "PEEK") OrElse

           material.Contains(
            "POM") OrElse

           material.Contains(
            "DELRIN") OrElse

           material.Contains(
            "NBR") OrElse

           material.Contains(
            "EPDM") OrElse

           material.Contains(
            "VITON") OrElse

           material.Contains(
            "FKM") OrElse

           material.Contains(
            "POLYAMIDE") OrElse

           material.Contains(
            "PA66") OrElse

           material.Contains(
            "PA 66") OrElse

           material.Contains(
            "PA6") OrElse

           material.Contains(
            "PA 6") Then

            Return "RUBBER/PLASTIC"

        End If


        '======================================================================
        ' REFRACTORY
        '======================================================================

        If material.Contains(
            "REFRACTORY") OrElse

           material.Contains(
            "CERAMIC") OrElse

           material.Contains(
            "FIREBRICK") OrElse

           material.Contains(
            "FIRE BRICK") Then

            Return "REFRACTORY"

        End If


        '======================================================================
        ' MARBLE / SLATE / GLASS
        '======================================================================

        If material.Contains(
            "MARBLE") OrElse

           material.Contains(
            "SLATE") OrElse

           material.Contains(
            "GLASS") Then

            Return "MARBLE/SLATE/GLASS"

        End If


        '======================================================================
        ' LEAD / ZINC / TIN / BABBITT
        '======================================================================

        If material.Contains(
            "LEAD") OrElse

           material.Contains(
            "ZINC") OrElse

           material.Contains(
            "BABBITT") OrElse

           material.Contains(
            "WHITE METAL") OrElse

           material.Contains(
            "TIN ALLOY") Then

            Return "LEAD"

        End If


        '======================================================================
        ' STEEL
        '======================================================================

        If material.Contains(
            "STEEL") OrElse

           material.Contains(
            "AISI") OrElse

           material.Contains(
            "SAE") OrElse

           material.Contains(
            "20MNCR5") OrElse

           material.Contains(
            "16MNCR5") OrElse

           material.Contains(
            "18CRNIMO") OrElse

           material.Contains(
            "42CRMO") OrElse

           material.Contains(
            "8620") OrElse

           material.Contains(
            "4140") OrElse

           material.Contains(
            "4340") OrElse

           material.Contains(
            "1045") OrElse

           material.Contains(
            "EN8") OrElse

           material.Contains(
            "EN 8") OrElse

           material.Contains(
            "EN19") OrElse

           material.Contains(
            "EN 19") OrElse

           material.Contains(
            "EN24") OrElse

           material.Contains(
            "EN 24") OrElse

           material.Contains(
            "C45") OrElse

           material.Contains(
            "S45C") OrElse

           material.Contains(
            "SCM") OrElse

           material.Contains(
            "CRMO") OrElse

           material.Contains(
            "CARBON STEEL") OrElse

           material.Contains(
            "ALLOY STEEL") Then

            Return "STEEL"

        End If


        '======================================================================
        ' GENERIC IRON
        '======================================================================

        If material.Contains(
            "IRON") Then

            Return "IRON"

        End If


        '======================================================================
        ' UNKNOWN
        '======================================================================

        If SKIP_UNKNOWN_MATERIAL Then

            Return ""

        End If


        Return "IRON"

    End Function


    '==========================================================================
    ' APPLY HATCH RULE
    '==========================================================================

    Private Sub ApplyHatchRule(
        ByVal hatch As Hatch,
        ByVal rule As HatchRule)


        Dim hatchBuilder As HatchBuilder = Nothing


        Try

            hatchBuilder =
                workPart.Annotations.Hatches.
                CreateHatchBuilder(
                    hatch)


            '------------------------------------------------------------------
            ' Hatch type
            '------------------------------------------------------------------

            hatchBuilder.AnnotationType =
                HatchBuilder.AnnotationTypes.Crosshatch


            '------------------------------------------------------------------
            ' Hatch fill settings
            '------------------------------------------------------------------

            Dim settings As HatchFillSettingsBuilder =
                hatchBuilder.HatchFillSettings


            '------------------------------------------------------------------
            ' Try to use Siemens default xhatch.chx
            '------------------------------------------------------------------

            Dim crossHatchFile As String =
                FindCrossHatchFile()


            If Not String.IsNullOrWhiteSpace(
                crossHatchFile) Then

                Try

                    settings.CrosshatchFile =
                        crossHatchFile

                Catch

                End Try

            End If


            '------------------------------------------------------------------
            ' Material pattern
            '------------------------------------------------------------------

            settings.Pattern =
                rule.Pattern


            settings.Angle =
                rule.Angle


            settings.Distance =
                rule.Distance


            settings.Scale =
                rule.Scale


            '------------------------------------------------------------------
            ' Commit
            '------------------------------------------------------------------

            hatchBuilder.Commit()


        Finally

            If hatchBuilder IsNot Nothing Then

                Try

                    hatchBuilder.Destroy()

                Catch

                End Try

            End If

        End Try

    End Sub


    '==========================================================================
    ' FIND SIEMENS NX XHATCH FILE
    '==========================================================================

    Private Function FindCrossHatchFile() As String


        Dim possibleFiles As New List(Of String)


        '----------------------------------------------------------------------
        ' UGII_BASE_DIR
        '----------------------------------------------------------------------

        Dim baseDir As String =
            Environment.GetEnvironmentVariable(
                "UGII_BASE_DIR")


        If Not String.IsNullOrWhiteSpace(
            baseDir) Then


            possibleFiles.Add(
                Path.Combine(
                    baseDir,
                    "UGII",
                    "xhatch.chx"))


            possibleFiles.Add(
                Path.Combine(
                    baseDir,
                    "xhatch.chx"))

        End If


        '----------------------------------------------------------------------
        ' UGII_ROOT_DIR
        '----------------------------------------------------------------------

        Dim rootDir As String =
            Environment.GetEnvironmentVariable(
                "UGII_ROOT_DIR")


        If Not String.IsNullOrWhiteSpace(
            rootDir) Then


            possibleFiles.Add(
                Path.Combine(
                    rootDir,
                    "xhatch.chx"))


            possibleFiles.Add(
                Path.Combine(
                    rootDir,
                    "UGII",
                    "xhatch.chx"))

        End If


        '----------------------------------------------------------------------
        ' Return first existing file
        '----------------------------------------------------------------------

        For Each fileName As String _
            In possibleFiles


            Try

                If File.Exists(
                    fileName) Then

                    Return fileName

                End If

            Catch

            End Try

        Next


        'If nothing found, NX will use the current/default
        'crosshatch configuration.
        Return ""

    End Function


    '==========================================================================
    ' SUMMARY
    '==========================================================================

    Private Sub PrintSummary()


        lw.WriteLine("")

        lw.WriteLine(
            "============================================================")

        lw.WriteLine(
            "                         SUMMARY")

        lw.WriteLine(
            "============================================================")


        lw.WriteLine(
            "Total hatches   : " &
            totalHatches.ToString())


        lw.WriteLine(
            "Updated hatches : " &
            changedHatches.ToString())


        lw.WriteLine(
            "Skipped hatches : " &
            skippedHatches.ToString())


        lw.WriteLine(
            "Errors           : " &
            errorHatches.ToString())


        lw.WriteLine(
            "============================================================")

        lw.WriteLine("")

    End Sub


    '==========================================================================
    ' UNLOAD
    '==========================================================================

    Public Function GetUnloadOption(
        ByVal dummy As String) As Integer


        Return Session.LibraryUnloadOption.Immediately

    End Function


End Module
