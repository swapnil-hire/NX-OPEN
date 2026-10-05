Option Strict On

Imports System
Imports System.IO
Imports System.Text
Imports System.Text.RegularExpressions
Imports System.Collections
Imports System.Collections.Generic
Imports System.Reflection

Imports NXOpen
Imports NXOpen.Annotations
Imports NXOpen.Drawings

Module NX2306_Drawing_Quality_Checker

    Private theSession As Session
    Private workPart As Part
    Private lw As ListingWindow

    Private ReadOnly Results As New List(Of CheckResult)

    '========================================================
    ' USER SETTINGS
    '========================================================

    'If True, normal dimensions without an explicit tolerance
    'are reported as WARNING. Keep False when your drawing uses
    'a general tolerance note/title-block tolerance.
    Private Const REQUIRE_EXPLICIT_TOLERANCE As Boolean = False

    'ISO sheet size comparison tolerance in mm.
    Private Const SHEET_SIZE_TOL_MM As Double = 2.0

    '========================================================
    ' DRAWING TEXT STANDARD
    '========================================================

    'Required font for dimensions and GD&T.
    Private Const REQUIRED_TEXT_FONT As String = "Verdana"

    'Allowed text heights in mm.
    'Current requirement: normal text 3.5 mm or large text 7.0 mm.
    Private Const ALLOWED_TEXT_HEIGHT_1 As Double = 3.5
    Private Const ALLOWED_TEXT_HEIGHT_2 As Double = 7.0

    'Tolerance used when comparing NX text height values.
    Private Const TEXT_HEIGHT_TOL_MM As Double = 0.02

    'If True, tolerance text on dimensions is checked separately.
    Private Const CHECK_DIMENSION_TOLERANCE_TEXT_STYLE As Boolean = True

    'If True, appended dimension text is checked separately.
    Private Const CHECK_DIMENSION_APPENDED_TEXT_STYLE As Boolean = True

    'Spell checking:
    'General dictionary should contain one word per line.
    'Engineering dictionary is optional and can contain your
    'company / automotive / NX terminology.
    Private Const GENERAL_DICTIONARY_PATH As String =
        "C:\NXTools\Dictionary\english_words.txt"

    Private Const ENGINEERING_DICTIONARY_PATH As String =
        "C:\NXTools\Dictionary\engineering_dictionary.txt"

    'Optional CSV report.
    Private Const EXPORT_CSV As Boolean = True

    '========================================================
    ' TITLE BLOCK / TEAMCENTER ATTRIBUTE ALIASES
    ' Edit these to match your environment.
    '========================================================

    Private ReadOnly REQUIRED_ATTRIBUTES As New Dictionary(Of String, String()) From {
        {"PART NUMBER", New String() {
            "DB_PART_NO", "DB_PART_NUMBER", "PART_NUMBER",
            "PART NO", "PART_NO", "ITEM_ID"
        }},
        {"REVISION", New String() {
            "DB_PART_REV", "DB_REVISION", "REVISION",
            "REV", "ITEM_REVISION"
        }},
        {"DESCRIPTION", New String() {
            "DB_PART_NAME", "PART_NAME", "DESCRIPTION",
            "PART DESCRIPTION", "PART_DESCRIPTION"
        }},
        {"MATERIAL", New String() {
            "MATERIAL", "MATERIAL_SPEC", "MATERIAL SPECIFICATION",
            "MATERIAL_SPECIFICATION"
        }},
        {"DESIGNER", New String() {
            "DESIGNER", "DRAWN_BY", "DRAWN BY", "AUTHOR"
        }}
    }

    '========================================================
    ' SPELL CHECK DATA
    '========================================================

    Private ReadOnly SpellDictionary As New HashSet(Of String)(
        StringComparer.OrdinalIgnoreCase)

    Private ReadOnly IgnoreWords As New HashSet(Of String)(
        StringComparer.OrdinalIgnoreCase) From {

        "NX", "CAD", "PMI", "GD", "GDT", "FCF",
        "TYP", "REF", "MAX", "MIN", "NTS",
        "LH", "RH", "OD", "ID", "PCD",
        "THRU", "THRO", "EQ", "EQL",
        "MATL", "QTY", "ASSY", "DWG", "REV", "DIM",
        "RFS", "MMC", "LMC",
        "BOM", "LH", "RH", "TIR",
        "DIA", "RAD", "RADIUS",
        "UNC", "UNF", "UNEF", "ISO",
        "LH", "RH", "TBD", "NA", "NIL",
        "AF", "ACROSS", "FLATS",
        "CBORE", "CSINK", "SFACE",
        "DEG", "MM", "CM", "M", "IN", "INCH",
        "KG", "GM", "NM", "MPA", "GPA",
        "HRC", "HV", "HB",
        "RA", "RZ",
        "SPC", "PPAP", "APQP",
        "OEM", "EV", "ICE",
        "CAD", "CAE", "CAM", "PLM", "TC", "JT", "STEP",
        "RMS", "RMR", "FOS", "BSC", "BASIC",
        "CRS", "CL", "SYM", "SIM", "ALT",
        "NO", "NOS", "REQD", "REQ", "SPEC", "STD",
        "MFG", "MACH", "HT", "HRS", "MIN", "MAX",
        "ECO", "ECN", "ECM", "VAVE", "DFMEA", "PFMEA",
        "RH", "LH", "FR", "RR",
        "PITCH", "ROOT", "MAJOR", "MINOR"
    }

    'Small built-in engineering vocabulary.
    'A full English word-list is still recommended.
    Private ReadOnly BuiltInWords As String() = {
        "assembly", "assemblies", "component", "components",
        "drawing", "drawings", "dimension", "dimensions",
        "tolerance", "tolerances", "datum", "datums",
        "surface", "finish", "finishing",
        "machining", "machine", "machined",
        "manufacture", "manufacturing", "manufacturer",
        "material", "materials",
        "heat", "treatment", "treated",
        "harden", "hardened", "hardening",
        "case", "carburized", "carburizing",
        "induction", "tempered", "annealed",
        "forging", "forged", "casting", "cast",
        "driveshaft", "propeller", "shaft", "axle",
        "spline", "splines", "splined",
        "gear", "gears", "bearing", "bearings",
        "housing", "carrier", "flange", "yoke",
        "runout", "concentricity", "coaxiality",
        "cylindricity", "flatness", "straightness",
        "parallelism", "perpendicularity", "position",
        "profile", "circularity", "symmetry",
        "counterbore", "countersink", "spotface",
        "thread", "threads", "threaded", "threadlocker",
        "diameter", "radius", "chamfer",
        "break", "sharp", "edges", "edge",
        "remove", "burrs", "burr",
        "deburr", "clean", "cleaning",
        "paint", "painted", "coating", "coated",
        "phosphate", "plated", "plating",
        "zinc", "nickel", "chrome",
        "weld", "welding", "welded",
        "inspection", "inspect", "checked",
        "reference", "typical",
        "unless", "otherwise", "specified",
        "apply", "applied", "required",
        "before", "after", "during",
        "shall", "must", "may",
        "all", "any", "and", "or", "of", "to", "for",
        "with", "without", "from", "on", "in", "at",
        "this", "that", "these", "those",
        "part", "parts", "number", "revision",
        "scale", "sheet", "sheets",
        "view", "views", "section", "detail",
        "note", "notes", "general",
        "length", "width", "height", "depth",
        "upper", "lower", "limit", "limits",
        "minimum", "maximum",
        "hole", "holes", "slot", "slots",
        "center", "centre", "centerline", "centreline",
        "mark", "marked", "marking",
        "protect", "protected", "protection",
        "pack", "packing", "package",
        "customer", "supplier",
        "drawing", "approval", "approved",
        "design", "designer",
        "weight", "mass",
        "do", "not", "use", "as", "shown"
    }

    Private SpellDictionaryReady As Boolean = False
    Private SpellWordsChecked As Integer = 0
    Private SpellErrorsFound As Integer = 0

    '========================================================
    ' RESULT CLASS
    '========================================================

    Private Class CheckResult
        Public Severity As String
        Public Category As String
        Public ObjectId As String
        Public Message As String

        Public Sub New(
            ByVal severityValue As String,
            ByVal categoryValue As String,
            ByVal objectIdValue As String,
            ByVal messageValue As String)

            Severity = severityValue
            Category = categoryValue
            ObjectId = objectIdValue
            Message = messageValue
        End Sub
    End Class

    '========================================================
    ' MAIN
    '========================================================

    Public Sub Main()

        theSession = Session.GetSession()
        workPart = theSession.Parts.Work
        lw = theSession.ListingWindow
        lw.Open()

        lw.WriteLine("")
        lw.WriteLine("============================================================")
        lw.WriteLine("    NX 2306 DRAWING QUALITY CHECKER - ENHANCED")
        lw.WriteLine("============================================================")
        lw.WriteLine("")

        If workPart Is Nothing Then
            lw.WriteLine("ERROR: No work part is open.")
            Return
        End If

        Try
            lw.WriteLine("PART: " & workPart.Leaf)
            lw.WriteLine("")

            LoadSpellDictionaries()

            CheckTitleBlockAndPartAttributes()
            CheckDrawingSheetsAndViews()
            CheckDimensions()
            CheckGDAndT()

            'Company drafting standard:
            'Verdana font and text height 3.5 mm or 7.0 mm.
            CheckDimensionFontAndHeight()
            CheckGdtFontAndHeight()

            CheckAnnotationInventory()
            CheckSpelling()

            PrintReport()

            If EXPORT_CSV Then
                ExportCsvReport()
            End If

        Catch ex As Exception
            lw.WriteLine("")
            lw.WriteLine("FATAL ERROR")
            lw.WriteLine(ex.Message)
            lw.WriteLine(ex.StackTrace)
        End Try

    End Sub

    '========================================================
    ' TITLE BLOCK + PART ATTRIBUTES
    '========================================================

    Private Sub CheckTitleBlockAndPartAttributes()

        'Check actual NX title-block objects.
        Try
            Dim titleBlocks() As TitleBlock =
                workPart.DraftingManager.TitleBlocks.ToArray()

            If titleBlocks Is Nothing OrElse titleBlocks.Length = 0 Then
                AddResult("WARNING", "TITLE BLOCK", "",
                          "No NX TitleBlock object found. The drawing may use a legacy tabular-note title block.")
            Else
                AddResult("PASS", "TITLE BLOCK", "",
                          titleBlocks.Length.ToString() & " NX title block object(s) found.")

                For Each tb As TitleBlock In titleBlocks
                    Try
                        If tb.Suppressed Then
                            AddResult("WARNING", "TITLE BLOCK",
                                      "Tag " & tb.Tag.ToString(),
                                      "Title block is suppressed.")
                        End If

                        If tb.IsOutOfDate Then
                            AddResult("WARNING", "TITLE BLOCK",
                                      "Tag " & tb.Tag.ToString(),
                                      "Title block is out of date.")
                        End If
                    Catch ex As Exception
                        AddResult("WARNING", "TITLE BLOCK",
                                  "Tag " & tb.Tag.ToString(),
                                  "Unable to inspect title block state: " & ex.Message)
                    End Try
                Next
            End If

        Catch ex As Exception
            AddResult("WARNING", "TITLE BLOCK", "",
                      "Unable to enumerate NX title blocks: " & ex.Message)
        End Try

        'Check attributes feeding the title block / Teamcenter.
        Dim attributes() As NXObject.AttributeInformation

        Try
            attributes = workPart.GetUserAttributes()
        Catch ex As Exception
            AddResult("ERROR", "TITLE BLOCK", workPart.Leaf,
                      "Unable to read part attributes: " & ex.Message)
            Return
        End Try

        For Each req As KeyValuePair(Of String, String()) In REQUIRED_ATTRIBUTES

            Dim logicalName As String = req.Key
            Dim aliases() As String = req.Value

            Dim found As Boolean = False
            Dim foundTitle As String = ""
            Dim foundValue As String = ""

            For Each att As NXObject.AttributeInformation In attributes

                If att.Unset Then Continue For

                For Each aliasName As String In aliases

                    If String.Equals(
                        att.Title.Trim(),
                        aliasName.Trim(),
                        StringComparison.OrdinalIgnoreCase) Then

                        found = True
                        foundTitle = att.Title
                        foundValue = AttributeValueAsString(att)
                        Exit For
                    End If
                Next

                If found Then Exit For
            Next

            If Not found Then
                AddResult("ERROR", "TITLE BLOCK / ATTRIBUTE",
                          logicalName,
                          "Required attribute not found.")

            ElseIf String.IsNullOrWhiteSpace(foundValue) Then
                AddResult("ERROR", "TITLE BLOCK / ATTRIBUTE",
                          logicalName,
                          "Attribute '" & foundTitle & "' exists but its value is blank.")

            Else
                AddResult("PASS", "TITLE BLOCK / ATTRIBUTE",
                          logicalName,
                          foundValue)
            End If

        Next

    End Sub

    Private Function AttributeValueAsString(
        ByVal info As NXObject.AttributeInformation) As String

        Try
            If info.Type = NXObject.AttributeType.String Then
                Return info.StringValue
            End If
        Catch
        End Try

        Try
            Return info.ToString()
        Catch
            Return ""
        End Try

    End Function

    '========================================================
    ' SHEETS + VIEWS
    '========================================================

    Private Sub CheckDrawingSheetsAndViews()

        Dim sheets() As DrawingSheet

        Try
            sheets = workPart.DrawingSheets.ToArray()
        Catch ex As Exception
            AddResult("ERROR", "SHEET", "",
                      "Unable to read drawing sheets: " & ex.Message)
            Return
        End Try

        If sheets Is Nothing OrElse sheets.Length = 0 Then
            AddResult("ERROR", "SHEET", "", "No drawing sheets found.")
            Return
        End If

        AddResult("PASS", "SHEET", "",
                  sheets.Length.ToString() & " drawing sheet(s) found.")

        Dim sheetIndex As Integer = 0

        For Each sheet As DrawingSheet In sheets

            sheetIndex += 1

            Try
                Dim sheetName As String = sheet.Name
                If String.IsNullOrWhiteSpace(sheetName) Then
                    sheetName = "Sheet " & sheetIndex.ToString()
                End If

                'Sheet size
                Dim isoSize As String = GetISOSize(sheet)

                If isoSize = "CUSTOM" Then
                    AddResult("WARNING", "SHEET SIZE", sheetName,
                              "Custom/non-ISO size: " &
                              Math.Round(sheet.Length, 2).ToString() & " x " &
                              Math.Round(sheet.Height, 2).ToString() &
                              " [" & sheet.Units.ToString() & "]")
                Else
                    AddResult("PASS", "SHEET SIZE", sheetName,
                              isoSize & " | " &
                              Math.Round(sheet.Length, 2).ToString() & " x " &
                              Math.Round(sheet.Height, 2).ToString() &
                              " [" & sheet.Units.ToString() & "]")
                End If

                'Update state
                If sheet.IsOutOfDate Then
                    AddResult("WARNING", "SHEET", sheetName,
                              "Sheet is out of date.")
                Else
                    AddResult("PASS", "SHEET", sheetName,
                              "Sheet update status OK.")
                End If

                'Border and zones
                Try
                    If sheet.BordersAndZones Is Nothing Then
                        AddResult("WARNING", "BORDER / ZONES", sheetName,
                                  "No BordersAndZones object found.")
                    Else
                        AddResult("PASS", "BORDER / ZONES", sheetName,
                                  "BordersAndZones object found.")
                    End If
                Catch ex As Exception
                    AddResult("WARNING", "BORDER / ZONES", sheetName,
                              "Unable to check border/zones: " & ex.Message)
                End Try

                'Projection angle
                Try
                    AddResult("INFO", "PROJECTION", sheetName,
                              sheet.ProjectionAngle.ToString())
                Catch ex As Exception
                    AddResult("WARNING", "PROJECTION", sheetName,
                              "Unable to read projection angle: " & ex.Message)
                End Try

                'Scale
                Dim numerator As Double = 0.0
                Dim denominator As Double = 0.0
                sheet.GetScale(numerator, denominator)

                If denominator = 0.0 Then
                    AddResult("WARNING", "SCALE", sheetName,
                              "Invalid sheet scale denominator.")
                Else
                    AddResult("PASS", "SCALE", sheetName,
                              numerator.ToString() & ":" & denominator.ToString())
                End If

                CheckViews(sheet, sheetName)

            Catch ex As Exception
                AddResult("ERROR", "SHEET",
                          "Sheet " & sheetIndex.ToString(),
                          "Unable to inspect sheet: " & ex.Message)
            End Try

        Next

    End Sub

    Private Sub CheckViews(
        ByVal sheet As DrawingSheet,
        ByVal sheetName As String)

        Dim views() As DraftingView

        Try
            views = sheet.GetDraftingViews()
        Catch ex As Exception
            AddResult("ERROR", "VIEW", sheetName,
                      "Unable to read drafting views: " & ex.Message)
            Return
        End Try

        If views Is Nothing OrElse views.Length = 0 Then
            AddResult("WARNING", "VIEW", sheetName,
                      "No drafting views found.")
            Return
        End If

        AddResult("PASS", "VIEW", sheetName,
                  views.Length.ToString() & " drafting view(s) found.")

        For Each viewObj As DraftingView In views

            Try
                Dim viewName As String = viewObj.Name
                If String.IsNullOrWhiteSpace(viewName) Then
                    viewName = "Tag " & viewObj.Tag.ToString()
                End If

                If viewObj.IsOutOfDate Then
                    AddResult("WARNING", "VIEW", viewName,
                              "View is out of date.")
                Else
                    AddResult("PASS", "VIEW", viewName,
                              "View update status OK.")
                End If

                If viewObj.IsBroken Then
                    AddResult("INFO", "VIEW", viewName,
                              "Broken view detected.")
                End If

                If viewObj.IsDecoration Then
                    AddResult("INFO", "VIEW", viewName,
                              "Decoration view.")
                End If

            Catch ex As Exception
                AddResult("ERROR", "VIEW",
                          "Tag " & viewObj.Tag.ToString(),
                          "Unable to inspect view: " & ex.Message)
            End Try
        Next

    End Sub

    '========================================================
    ' DIMENSIONS
    '========================================================

    Private Sub CheckDimensions()

        Dim dimensions() As Dimension

        Try
            dimensions = workPart.Dimensions.ToArray()
        Catch ex As Exception
            AddResult("ERROR", "DIMENSION", "",
                      "Unable to read dimensions: " & ex.Message)
            Return
        End Try

        If dimensions Is Nothing OrElse dimensions.Length = 0 Then
            AddResult("WARNING", "DIMENSION", "",
                      "No dimensions found.")
            Return
        End If

        AddResult("PASS", "DIMENSION", "",
                  dimensions.Length.ToString() & " dimension(s) found.")

        Dim counter As Integer = 0

        For Each dimObj As Dimension In dimensions

            counter += 1

            Dim dimId As String =
                "DIM-" & counter.ToString() &
                " [Tag " & dimObj.Tag.ToString() & "]"

            Try
                If dimObj.Suppressed Then
                    AddResult("WARNING", "DIMENSION", dimId,
                              "Dimension is suppressed.")
                End If

                If dimObj.IsOutOfDate Then
                    AddResult("WARNING", "DIMENSION", dimId,
                              "Dimension is out of date.")
                End If

                If dimObj.NumberOfAssociativities = 0 Then
                    AddResult("WARNING", "DIMENSION", dimId,
                              "Dimension has no associativity.")
                Else
                    AddResult("PASS", "DIMENSION", dimId,
                              "Associativities = " &
                              dimObj.NumberOfAssociativities.ToString())
                End If

                If dimObj.ReferenceDimensionFlag Then
                    AddResult("INFO", "DIMENSION", dimId,
                              "Reference dimension.")
                End If

                If dimObj.InspectionDimensionFlag Then
                    AddResult("INFO", "DIMENSION", dimId,
                              "Inspection dimension.")
                End If

                AddResult("INFO", "DIMENSION PRECISION", dimId,
                          "Nominal decimals = " &
                          dimObj.NominalDecimalPlaces.ToString() &
                          ", tolerance decimals = " &
                          dimObj.ToleranceDecimalPlaces.ToString())

                Dim tolType As String = dimObj.ToleranceType.ToString()

                AddResult("INFO", "TOLERANCE", dimId,
                          "Type = " & tolType)

                If Not tolType.Equals(
                    "None",
                    StringComparison.OrdinalIgnoreCase) Then

                    AddResult("PASS", "TOLERANCE", dimId,
                              "Upper = " &
                              dimObj.UpperToleranceValue.ToString() &
                              ", Lower = " &
                              dimObj.LowerToleranceValue.ToString())

                ElseIf REQUIRE_EXPLICIT_TOLERANCE AndAlso
                       Not dimObj.ReferenceDimensionFlag Then

                    AddResult("WARNING", "TOLERANCE", dimId,
                              "No explicit tolerance applied.")
                End If

                Dim mainText() As String = Nothing
                Dim dualText() As String = Nothing
                dimObj.GetDimensionText(mainText, dualText)

                If mainText Is Nothing OrElse mainText.Length = 0 Then
                    AddResult("WARNING", "DIMENSION TEXT", dimId,
                              "Displayed dimension text is empty.")
                End If

                Dim views() As View = dimObj.GetViews()

                If views Is Nothing OrElse views.Length = 0 Then
                    AddResult("INFO", "DIMENSION", dimId,
                              "No explicit annotation-view dependency.")
                End If

            Catch ex As Exception
                AddResult("ERROR", "DIMENSION", dimId,
                          "Unable to inspect dimension: " & ex.Message)
            End Try
        Next

    End Sub

    '========================================================
    ' GD&T
    '========================================================

    Private Sub CheckGDAndT()

        Dim total As Integer = 0
        Dim fcfCount As Integer = 0

        Try
            For Each gdtObj As Gdt In workPart.Gdts

                total += 1

                Dim objId As String =
                    "GDT-" & total.ToString() &
                    " [Tag " & gdtObj.Tag.ToString() & "]"

                Try
                    If gdtObj.Suppressed Then
                        AddResult("WARNING", "GD&T", objId,
                                  "GD&T object is suppressed.")
                    End If

                    If gdtObj.IsOutOfDate Then
                        AddResult("WARNING", "GD&T", objId,
                                  "GD&T object is out of date.")
                    End If

                    If gdtObj.NumberOfAssociativities = 0 Then
                        AddResult("WARNING", "GD&T", objId,
                                  "No associativity detected.")
                    Else
                        AddResult("PASS", "GD&T", objId,
                                  "Associativities = " &
                                  gdtObj.NumberOfAssociativities.ToString())
                    End If

                    Dim textLines() As String = ExtractTextLines(gdtObj)

                    If textLines Is Nothing OrElse textLines.Length = 0 Then
                        AddResult("WARNING", "GD&T", objId,
                                  "No readable text returned.")
                    Else
                        AddResult("PASS", "GD&T", objId,
                                  String.Join(" | ", textLines))
                    End If

                    Try
                        Dim frameData() As FcfFrameData =
                            gdtObj.GetFcfFrameDataArray()

                        If frameData IsNot Nothing AndAlso
                           frameData.Length > 0 Then

                            fcfCount += 1

                            For Each frame As FcfFrameData In frameData
                                Try
                                    AddResult("INFO", "FCF", objId,
                                              "Characteristic = " &
                                              frame.GeometricCharacteristic.ToString())
                                Catch
                                Finally
                                    Try
                                        frame.Dispose()
                                    Catch
                                    End Try
                                End Try
                            Next
                        End If
                    Catch
                        'Not every GDT object is an FCF.
                    End Try

                Catch ex As Exception
                    AddResult("ERROR", "GD&T", objId,
                              "Unable to inspect GD&T object: " & ex.Message)
                End Try
            Next

            If total = 0 Then
                AddResult("INFO", "GD&T", "",
                          "No drafting GD&T objects found.")
            Else
                AddResult("PASS", "GD&T", "",
                          "Total GD&T objects = " & total.ToString() &
                          ", FCF objects detected = " & fcfCount.ToString())
            End If

        Catch ex As Exception
            AddResult("ERROR", "GD&T", "",
                      "Unable to enumerate GD&T objects: " & ex.Message)
        End Try

    End Sub

    '========================================================
    ' DIMENSION FONT + HEIGHT STANDARD CHECK
    '
    ' Required:
    '   Font   = Verdana
    '   Height = 3.5 mm OR 7.0 mm
    '
    ' Checks:
    '   - Dimension text
    '   - Tolerance text
    '   - Appended text
    '========================================================

    Private Sub CheckDimensionFontAndHeight()

        Dim dimensions() As Dimension

        Try
            dimensions = workPart.Dimensions.ToArray()
        Catch ex As Exception
            AddResult("ERROR", "DIMENSION STYLE", "",
                      "Unable to enumerate dimensions for style checking: " &
                      ex.Message)
            Return
        End Try

        If dimensions Is Nothing OrElse dimensions.Length = 0 Then
            Return
        End If

        Dim counter As Integer = 0

        For Each dimObj As Dimension In dimensions

            counter += 1

            Dim objectId As String =
                "DIM-" & counter.ToString() &
                " [Tag " & dimObj.Tag.ToString() & "]"

            Dim prefs As LetteringPreferences = Nothing

            Try
                prefs = dimObj.GetLetteringPreferences()

                'Nominal / primary dimension text.
                Dim dimText As Lettering =
                    prefs.GetDimensionText()

                CheckLetteringStandard(
                    dimText,
                    "DIMENSION FONT/HEIGHT",
                    objectId,
                    "Dimension text")

                'Tolerance text can have an independent lettering style.
                If CHECK_DIMENSION_TOLERANCE_TEXT_STYLE Then

                    Dim toleranceText As Lettering =
                        prefs.GetToleranceText()

                    CheckLetteringStandard(
                        toleranceText,
                        "DIMENSION TOLERANCE STYLE",
                        objectId,
                        "Tolerance text")

                End If

                'Prefix/suffix/appended text can also have its own style.
                If CHECK_DIMENSION_APPENDED_TEXT_STYLE Then

                    Dim appendedText As Lettering =
                        prefs.GetAppendedText()

                    CheckLetteringStandard(
                        appendedText,
                        "DIMENSION APPENDED TEXT STYLE",
                        objectId,
                        "Appended text")

                End If

            Catch ex As Exception

                AddResult(
                    "ERROR",
                    "DIMENSION STYLE",
                    objectId,
                    "Unable to check lettering preferences: " &
                    ex.Message)

            Finally

                If prefs IsNot Nothing Then
                    Try
                        prefs.Dispose()
                    Catch
                    End Try
                End If

            End Try

        Next

    End Sub


    '========================================================
    ' GD&T FONT + HEIGHT STANDARD CHECK
    '
    ' GD&T uses general lettering preferences.
    ' Required:
    '   Font   = Verdana
    '   Height = 3.5 mm OR 7.0 mm
    '========================================================

    Private Sub CheckGdtFontAndHeight()

        Dim counter As Integer = 0

        Try

            For Each gdtObj As Gdt In workPart.Gdts

                counter += 1

                Dim objectId As String =
                    "GDT-" & counter.ToString() &
                    " [Tag " & gdtObj.Tag.ToString() & "]"

                Dim prefs As LetteringPreferences = Nothing

                Try
                    prefs = gdtObj.GetLetteringPreferences()

                    Dim generalText As Lettering =
                        prefs.GetGeneralText()

                    CheckLetteringStandard(
                        generalText,
                        "GD&T FONT/HEIGHT",
                        objectId,
                        "GD&T text")

                    'Useful diagnostic because GD&T frame height is
                    'affected by the lettering preferences.
                    Try
                        AddResult(
                            "INFO",
                            "GD&T FRAME",
                            objectId,
                            "GDT frame-height factor = " &
                            prefs.GdtFrameHeightFactor.ToString("0.###"))
                    Catch
                    End Try

                Catch ex As Exception

                    AddResult(
                        "ERROR",
                        "GD&T STYLE",
                        objectId,
                        "Unable to check GD&T lettering preferences: " &
                        ex.Message)

                Finally

                    If prefs IsNot Nothing Then
                        Try
                            prefs.Dispose()
                        Catch
                        End Try
                    End If

                End Try

            Next

        Catch ex As Exception

            AddResult(
                "ERROR",
                "GD&T STYLE",
                "",
                "Unable to enumerate GD&T for style checking: " &
                ex.Message)

        End Try

    End Sub


    '========================================================
    ' COMMON LETTERING STANDARD CHECK
    '========================================================

    Private Sub CheckLetteringStandard(
        ByVal letteringData As Lettering,
        ByVal category As String,
        ByVal objectId As String,
        ByVal textRole As String)

        Dim fontId As Integer = -1
        Dim fontName As String = ""
        Dim textHeight As Double = 0.0

        Try
            fontId = letteringData.Cfw.Font
        Catch ex As Exception
            AddResult(
                "ERROR",
                category,
                objectId,
                textRole &
                ": unable to read font ID: " &
                ex.Message)
            Return
        End Try

        Try
            fontName = GetNxFontName(fontId)
        Catch ex As Exception
            fontName = "Font ID " & fontId.ToString()
        End Try

        Try
            textHeight = letteringData.Size
        Catch ex As Exception
            AddResult(
                "ERROR",
                category,
                objectId,
                textRole &
                ": unable to read text height: " &
                ex.Message)
            Return
        End Try


        '------------------------------------------------------
        ' FONT
        '------------------------------------------------------

        If String.Equals(
            fontName.Trim(),
            REQUIRED_TEXT_FONT,
            StringComparison.OrdinalIgnoreCase) Then

            AddResult(
                "PASS",
                category,
                objectId,
                textRole &
                " | Font = " &
                fontName)

        Else

            AddResult(
                "ERROR",
                category,
                objectId,
                textRole &
                " | Font = " &
                fontName &
                " | Required = " &
                REQUIRED_TEXT_FONT)

        End If


        '------------------------------------------------------
        ' HEIGHT
        '------------------------------------------------------

        If IsAllowedTextHeight(textHeight) Then

            AddResult(
                "PASS",
                category,
                objectId,
                textRole &
                " | Height = " &
                textHeight.ToString("0.###") &
                " mm")

        Else

            AddResult(
                "ERROR",
                category,
                objectId,
                textRole &
                " | Height = " &
                textHeight.ToString("0.###") &
                " mm | Required = " &
                ALLOWED_TEXT_HEIGHT_1.ToString("0.###") &
                " or " &
                ALLOWED_TEXT_HEIGHT_2.ToString("0.###") &
                " mm")

        End If

    End Sub


    Private Function IsAllowedTextHeight(
        ByVal value As Double) As Boolean

        Return
            Math.Abs(value - ALLOWED_TEXT_HEIGHT_1) <= TEXT_HEIGHT_TOL_MM OrElse
            Math.Abs(value - ALLOWED_TEXT_HEIGHT_2) <= TEXT_HEIGHT_TOL_MM

    End Function


    '========================================================
    ' NX FONT NAME LOOKUP
    '
    ' Uses reflection so the journal remains tolerant of minor
    ' FontCollection API differences across NX installations.
    '========================================================

    Private Function GetNxFontName(
        ByVal fontId As Integer) As String

        Try
            Dim fontCollection As Object =
                workPart.Fonts

            Dim fontCollectionType As Type =
                fontCollection.GetType()

            Dim getFontNameMethod As MethodInfo =
                fontCollectionType.GetMethod(
                    "GetFontName",
                    New Type() {GetType(Integer)})

            If getFontNameMethod IsNot Nothing Then

                Dim value As Object =
                    getFontNameMethod.Invoke(
                        fontCollection,
                        New Object() {fontId})

                If value IsNot Nothing Then
                    Return value.ToString()
                End If

            End If

        Catch
        End Try

        Return "Font ID " & fontId.ToString()

    End Function


    '========================================================
    ' ANNOTATION INVENTORY
    ' Counts useful drafting objects.
    '========================================================

    Private Sub CheckAnnotationInventory()

        Try
            ReportCollectionCount("CENTERLINE",
                                  workPart.Annotations.Centerlines)
        Catch ex As Exception
            AddResult("WARNING", "CENTERLINE", "",
                      "Unable to count centerlines: " & ex.Message)
        End Try

        Try
            ReportCollectionCount("SURFACE FINISH",
                                  workPart.Annotations.DraftingSurfaceFinishSymbols)
        Catch ex As Exception
            AddResult("WARNING", "SURFACE FINISH", "",
                      "Unable to count surface-finish symbols: " & ex.Message)
        End Try

        Try
            ReportCollectionCount("WELD SYMBOL",
                                  workPart.Annotations.Welds)
        Catch ex As Exception
            AddResult("WARNING", "WELD SYMBOL", "",
                      "Unable to count weld symbols: " & ex.Message)
        End Try

        Try
            ReportCollectionCount("DATUM",
                                  workPart.Annotations.Datums)
        Catch ex As Exception
            AddResult("WARNING", "DATUM", "",
                      "Unable to count datum symbols: " & ex.Message)
        End Try

        Try
            ReportCollectionCount("DATUM TARGET",
                                  workPart.Annotations.DatumTargets)
        Catch ex As Exception
            AddResult("WARNING", "DATUM TARGET", "",
                      "Unable to count datum targets: " & ex.Message)
        End Try

        Try
            ReportCollectionCount("ID / BALLOON SYMBOL",
                                  workPart.Annotations.IdSymbols)
        Catch ex As Exception
            AddResult("WARNING", "ID / BALLOON SYMBOL", "",
                      "Unable to count ID symbols: " & ex.Message)
        End Try

        Try
            ReportCollectionCount("HATCH",
                                  workPart.Annotations.Hatches)
        Catch ex As Exception
            AddResult("WARNING", "HATCH", "",
                      "Unable to count hatches: " & ex.Message)
        End Try

    End Sub

    Private Sub ReportCollectionCount(
        ByVal category As String,
        ByVal collectionObject As Object)

        Dim count As Integer = CountEnumerable(collectionObject)

        AddResult("INFO", category, "",
                  "Count = " & count.ToString())

    End Sub

    Private Function CountEnumerable(
        ByVal collectionObject As Object) As Integer

        Dim count As Integer = 0

        Dim enumerableObject As IEnumerable =
            TryCast(collectionObject, IEnumerable)

        If enumerableObject Is Nothing Then Return 0

        For Each obj As Object In enumerableObject
            count += 1
        Next

        Return count

    End Function

    '========================================================
    ' SPELL CHECK
    ' Checks:
    '   - Notes
    '   - Labels
    '   - Dimension displayed text
    '   - GD&T readable text
    '========================================================

    Private Sub LoadSpellDictionaries()

        SpellDictionary.Clear()

        For Each word As String In BuiltInWords
            AddDictionaryWord(word)
        Next

        Dim loadedGeneral As Integer =
            LoadDictionaryFile(GENERAL_DICTIONARY_PATH)

        Dim loadedEngineering As Integer =
            LoadDictionaryFile(ENGINEERING_DICTIONARY_PATH)

        If loadedGeneral = 0 Then
            AddResult("WARNING", "SPELLING", "",
                      "General dictionary not loaded: " &
                      GENERAL_DICTIONARY_PATH &
                      ". Built-in engineering vocabulary will still be used, " &
                      "but false spelling warnings may occur.")
        Else
            AddResult("PASS", "SPELLING", "",
                      "General dictionary loaded: " &
                      loadedGeneral.ToString() & " words.")
        End If

        If loadedEngineering > 0 Then
            AddResult("PASS", "SPELLING", "",
                      "Engineering dictionary loaded: " &
                      loadedEngineering.ToString() & " words.")
        Else
            AddResult("INFO", "SPELLING", "",
                      "Optional engineering dictionary not loaded: " &
                      ENGINEERING_DICTIONARY_PATH)
        End If

        SpellDictionaryReady = (SpellDictionary.Count > 0)

    End Sub

    Private Function LoadDictionaryFile(
        ByVal filePath As String) As Integer

        If String.IsNullOrWhiteSpace(filePath) Then Return 0
        If Not File.Exists(filePath) Then Return 0

        Dim added As Integer = 0

        Try
            For Each line As String In File.ReadAllLines(filePath)

                Dim word As String = line.Trim()

                If word = "" Then Continue For
                If word.StartsWith("#") Then Continue For

                If AddDictionaryWord(word) Then
                    added += 1
                End If
            Next
        Catch ex As Exception
            AddResult("WARNING", "SPELLING", "",
                      "Unable to read dictionary '" &
                      filePath & "': " & ex.Message)
        End Try

        Return added

    End Function

    Private Function AddDictionaryWord(
        ByVal rawWord As String) As Boolean

        If String.IsNullOrWhiteSpace(rawWord) Then Return False

        Dim word As String =
            rawWord.Trim().ToLowerInvariant()

        Return SpellDictionary.Add(word)

    End Function

    Private Sub CheckSpelling()

        If Not SpellDictionaryReady Then
            AddResult("WARNING", "SPELLING", "",
                      "Spell check skipped because no dictionary words are available.")
            Return
        End If

        SpellWordsChecked = 0
        SpellErrorsFound = 0

        Dim alreadyChecked As New HashSet(Of String)(
            StringComparer.OrdinalIgnoreCase)

        'Notes
        Try
            For Each baseNote As BaseNote In workPart.Notes.ToArray()

                Try
                    Dim lines() As String = ExtractTextLines(baseNote)

                    CheckTextLines(
                        lines,
                        "NOTE",
                        "Tag " & baseNote.Tag.ToString(),
                        alreadyChecked)

                Catch ex As Exception
                    AddResult("WARNING", "SPELLING",
                              "NOTE Tag " & baseNote.Tag.ToString(),
                              "Unable to read note text: " & ex.Message)
                End Try
            Next
        Catch ex As Exception
            AddResult("WARNING", "SPELLING", "NOTES",
                      "Unable to enumerate notes: " & ex.Message)
        End Try

        'Labels
        Try
            For Each labelObj As Label In workPart.Labels.ToArray()

                Try
                    Dim lines() As String = ExtractTextLines(labelObj)

                    CheckTextLines(
                        lines,
                        "LABEL",
                        "Tag " & labelObj.Tag.ToString(),
                        alreadyChecked)

                Catch ex As Exception
                    AddResult("WARNING", "SPELLING",
                              "LABEL Tag " & labelObj.Tag.ToString(),
                              "Unable to read label text: " & ex.Message)
                End Try
            Next
        Catch ex As Exception
            AddResult("WARNING", "SPELLING", "LABELS",
                      "Unable to enumerate labels: " & ex.Message)
        End Try

        'Dimension displayed text
        Try
            For Each dimObj As Dimension In workPart.Dimensions.ToArray()

                Try
                    Dim mainText() As String = Nothing
                    Dim dualText() As String = Nothing

                    dimObj.GetDimensionText(mainText, dualText)

                    CheckTextLines(
                        mainText,
                        "DIMENSION TEXT",
                        "Tag " & dimObj.Tag.ToString(),
                        alreadyChecked)

                    CheckTextLines(
                        dualText,
                        "DUAL DIMENSION TEXT",
                        "Tag " & dimObj.Tag.ToString(),
                        alreadyChecked)

                Catch ex As Exception
                    AddResult("WARNING", "SPELLING",
                              "DIMENSION Tag " & dimObj.Tag.ToString(),
                              "Unable to read dimension text: " & ex.Message)
                End Try
            Next
        Catch ex As Exception
            AddResult("WARNING", "SPELLING", "DIMENSIONS",
                      "Unable to enumerate dimension text: " & ex.Message)
        End Try

        'GD&T text
        Try
            For Each gdtObj As Gdt In workPart.Gdts

                Try
                    Dim lines() As String = ExtractTextLines(gdtObj)

                    CheckTextLines(
                        lines,
                        "GD&T TEXT",
                        "Tag " & gdtObj.Tag.ToString(),
                        alreadyChecked)

                Catch ex As Exception
                    AddResult("WARNING", "SPELLING",
                              "GD&T Tag " & gdtObj.Tag.ToString(),
                              "Unable to read GD&T text: " & ex.Message)
                End Try
            Next
        Catch ex As Exception
            AddResult("WARNING", "SPELLING", "GD&T",
                      "Unable to enumerate GD&T text: " & ex.Message)
        End Try

        If SpellErrorsFound = 0 Then
            AddResult("PASS", "SPELLING", "",
                      SpellWordsChecked.ToString() &
                      " unique word(s) checked; no possible spelling errors.")
        Else
            AddResult("WARNING", "SPELLING", "",
                      SpellWordsChecked.ToString() &
                      " unique word(s) checked; " &
                      SpellErrorsFound.ToString() &
                      " possible spelling error(s).")
        End If

    End Sub

    Private Sub CheckTextLines(
        ByVal lines() As String,
        ByVal sourceType As String,
        ByVal objectId As String,
        ByVal alreadyChecked As HashSet(Of String))

        If lines Is Nothing Then Return

        For Each line As String In lines
            CheckTextForSpelling(
                line,
                sourceType,
                objectId,
                alreadyChecked)
        Next

    End Sub

    Private Sub CheckTextForSpelling(
        ByVal inputText As String,
        ByVal sourceType As String,
        ByVal objectId As String,
        ByVal alreadyChecked As HashSet(Of String))

        If String.IsNullOrWhiteSpace(inputText) Then Return

        'Remove NX formatting sequences inside <...> where possible.
        Dim visibleText As String =
            Regex.Replace(inputText, "<[^>]*>", " ")

        'Alphabetic word tokens. Apostrophes/hyphens are retained.
        Dim matches As MatchCollection =
            Regex.Matches(
                visibleText,
                "[A-Za-z]+(?:['-][A-Za-z]+)*")

        For Each m As Match In matches

            Dim rawWord As String = m.Value

            If ShouldIgnoreWord(rawWord) Then Continue For

            Dim word As String =
                rawWord.ToLowerInvariant()

            If alreadyChecked.Contains(word) Then Continue For

            alreadyChecked.Add(word)
            SpellWordsChecked += 1

            If Not SpellDictionary.Contains(word) Then

                SpellErrorsFound += 1

                Dim suggestion As String =
                    FindSuggestion(word)

                Dim msg As String =
                    "Possible spelling error: '" &
                    rawWord & "'"

                If suggestion <> "" Then
                    msg &= " | Suggestion: '" &
                           suggestion.ToUpperInvariant() & "'"
                End If

                AddResult("WARNING", "SPELLING",
                          sourceType & " | " & objectId,
                          msg)
            End If
        Next

    End Sub

    Private Function ShouldIgnoreWord(
        ByVal rawWord As String) As Boolean

        If String.IsNullOrWhiteSpace(rawWord) Then Return True

        Dim word As String = rawWord.Trim()

        If word.Length <= 1 Then Return True

        If IgnoreWords.Contains(word) Then Return True

        'Roman numerals commonly used in drawings.
        If Regex.IsMatch(word, "^[IVXLCDM]+$",
                         RegexOptions.IgnoreCase) Then
            Return True
        End If

        'Short all-cap drawing abbreviations such as ABC, CRS, DIA etc.
        'Known words should still be added to the dictionary if you want
        'them checked. This rule intentionally reduces false positives.
        If word.Length <= 4 AndAlso
           Regex.IsMatch(word, "^[A-Z]+$") Then
            Return True
        End If

        Return False

    End Function

    Private Function FindSuggestion(
        ByVal misspelledWord As String) As String

        Dim bestWord As String = ""
        Dim bestDistance As Integer = 3
        Dim firstChar As Char = misspelledWord(0)

        For Each candidate As String In SpellDictionary

            If candidate.Length = 0 Then Continue For
            If candidate(0) <> firstChar Then Continue For

            If Math.Abs(
                candidate.Length -
                misspelledWord.Length) > 2 Then

                Continue For
            End If

            Dim distance As Integer =
                LevenshteinDistance(
                    misspelledWord,
                    candidate,
                    bestDistance)

            If distance < bestDistance Then
                bestDistance = distance
                bestWord = candidate

                If distance = 1 Then
                    'A distance of 1 is already a strong suggestion.
                End If
            End If
        Next

        If bestDistance <= 2 Then
            Return bestWord
        End If

        Return ""

    End Function

    Private Function LevenshteinDistance(
        ByVal source As String,
        ByVal target As String,
        ByVal stopAbove As Integer) As Integer

        Dim n As Integer = source.Length
        Dim m As Integer = target.Length

        If Math.Abs(n - m) > stopAbove Then
            Return stopAbove + 1
        End If

        Dim previous(m) As Integer
        Dim current(m) As Integer

        For j As Integer = 0 To m
            previous(j) = j
        Next

        For i As Integer = 1 To n

            current(0) = i
            Dim rowMin As Integer = current(0)

            For j As Integer = 1 To m

                Dim cost As Integer =
                    If(source(i - 1) = target(j - 1), 0, 1)

                current(j) =
                    Math.Min(
                        Math.Min(
                            current(j - 1) + 1,
                            previous(j) + 1),
                        previous(j - 1) + cost)

                If current(j) < rowMin Then
                    rowMin = current(j)
                End If
            Next

            If rowMin > stopAbove Then
                Return stopAbove + 1
            End If

            Dim temp() As Integer = previous
            previous = current
            current = temp
        Next

        Return previous(m)

    End Function

    '========================================================
    ' GENERIC GETTEXT USING REFLECTION
    ' This lets the same helper work for Note, Label and Gdt.
    '========================================================

    Private Function ExtractTextLines(
        ByVal obj As Object) As String()

        If obj Is Nothing Then
            Return New String() {}
        End If

        Try
            Dim methodInfo As MethodInfo =
                obj.GetType().GetMethod(
                    "GetText",
                    Type.EmptyTypes)

            If methodInfo Is Nothing Then
                Return New String() {}
            End If

            Dim result As Object =
                methodInfo.Invoke(obj, Nothing)

            Dim textArray() As String =
                TryCast(result, String())

            If textArray Is Nothing Then
                Return New String() {}
            End If

            Return textArray

        Catch
            Return New String() {}
        End Try

    End Function

    '========================================================
    ' ISO SHEET SIZE
    '========================================================

    Private Function GetISOSize(
        ByVal sheet As DrawingSheet) As String

        Dim width As Double = sheet.Length
        Dim height As Double = sheet.Height

        Dim unitName As String =
            sheet.Units.ToString().ToUpperInvariant()

        If unitName.Contains("INCH") OrElse
           unitName.Contains("ENGLISH") Then

            width *= 25.4
            height *= 25.4
        End If

        Dim shortSide As Double = Math.Min(width, height)
        Dim longSide As Double = Math.Max(width, height)

        If IsSize(shortSide, longSide, 841, 1189) Then Return "A0"
        If IsSize(shortSide, longSide, 594, 841) Then Return "A1"
        If IsSize(shortSide, longSide, 420, 594) Then Return "A2"
        If IsSize(shortSide, longSide, 297, 420) Then Return "A3"
        If IsSize(shortSide, longSide, 210, 297) Then Return "A4"

        Return "CUSTOM"

    End Function

    Private Function IsSize(
        ByVal actualShort As Double,
        ByVal actualLong As Double,
        ByVal expectedShort As Double,
        ByVal expectedLong As Double) As Boolean

        Return
            Math.Abs(actualShort - expectedShort) <= SHEET_SIZE_TOL_MM AndAlso
            Math.Abs(actualLong - expectedLong) <= SHEET_SIZE_TOL_MM

    End Function

    '========================================================
    ' RESULTS
    '========================================================

    Private Sub AddResult(
        ByVal severity As String,
        ByVal category As String,
        ByVal objectId As String,
        ByVal message As String)

        Results.Add(
            New CheckResult(
                severity,
                category,
                objectId,
                message))

    End Sub

    Private Sub PrintReport()

        lw.WriteLine("")
        lw.WriteLine("============================================================")
        lw.WriteLine("                    CHECK RESULTS")
        lw.WriteLine("============================================================")
        lw.WriteLine("")

        Dim passCount As Integer = 0
        Dim warningCount As Integer = 0
        Dim errorCount As Integer = 0
        Dim infoCount As Integer = 0

        For Each item As CheckResult In Results

            Select Case item.Severity.ToUpperInvariant()
                Case "PASS"
                    passCount += 1
                Case "WARNING"
                    warningCount += 1
                Case "ERROR"
                    errorCount += 1
                Case Else
                    infoCount += 1
            End Select

            Dim objectText As String = ""

            If Not String.IsNullOrWhiteSpace(item.ObjectId) Then
                objectText = " | " & item.ObjectId
            End If

            lw.WriteLine(
                "[" & item.Severity & "] " &
                item.Category &
                objectText)

            lw.WriteLine("    " & item.Message)
        Next

        lw.WriteLine("")
        lw.WriteLine("============================================================")
        lw.WriteLine("                       SUMMARY")
        lw.WriteLine("============================================================")
        lw.WriteLine("PASS     : " & passCount.ToString())
        lw.WriteLine("WARNING  : " & warningCount.ToString())
        lw.WriteLine("ERROR    : " & errorCount.ToString())
        lw.WriteLine("INFO     : " & infoCount.ToString())
        lw.WriteLine("------------------------------------------------------------")

        If errorCount = 0 AndAlso warningCount = 0 Then
            lw.WriteLine("DRAWING CHECK RESULT : PASS")
        ElseIf errorCount = 0 Then
            lw.WriteLine("DRAWING CHECK RESULT : PASS WITH WARNINGS")
        Else
            lw.WriteLine("DRAWING CHECK RESULT : REVIEW REQUIRED")
        End If

        lw.WriteLine("============================================================")
        lw.WriteLine("")

    End Sub

    '========================================================
    ' CSV REPORT
    '========================================================

    Private Sub ExportCsvReport()

        Try
            Dim outputFolder As String = ""

            Try
                If Not String.IsNullOrWhiteSpace(workPart.FullPath) Then
                    outputFolder =
                        Path.GetDirectoryName(workPart.FullPath)
                End If
            Catch
                outputFolder = ""
            End Try

            If String.IsNullOrWhiteSpace(outputFolder) OrElse
               Not Directory.Exists(outputFolder) Then

                outputFolder = Path.GetTempPath()
            End If

            Dim safePartName As String =
                MakeSafeFileName(workPart.Leaf)

            Dim csvPath As String =
                Path.Combine(
                    outputFolder,
                    safePartName &
                    "_Drawing_Check_" &
                    DateTime.Now.ToString("yyyyMMdd_HHmmss") &
                    ".csv")

            Using writer As New StreamWriter(
                csvPath,
                False,
                Encoding.UTF8)

                writer.WriteLine(
                    "Severity,Category,Object,Message")

                For Each item As CheckResult In Results

                    writer.WriteLine(
                        Csv(item.Severity) & "," &
                        Csv(item.Category) & "," &
                        Csv(item.ObjectId) & "," &
                        Csv(item.Message))
                Next
            End Using

            lw.WriteLine(
                "CSV REPORT: " & csvPath)

        Catch ex As Exception
            lw.WriteLine(
                "CSV export failed: " &
                ex.Message)
        End Try

    End Sub

    Private Function Csv(
        ByVal value As String) As String

        If value Is Nothing Then value = ""

        Return """" &
               value.Replace("""", """""") &
               """"

    End Function

    Private Function MakeSafeFileName(
        ByVal value As String) As String

        If String.IsNullOrWhiteSpace(value) Then
            Return "NX_Drawing"
        End If

        Dim result As String = value

        For Each ch As Char In Path.GetInvalidFileNameChars()
            result = result.Replace(ch, "_"c)
        Next

        Return result

    End Function

    '========================================================
    ' NX UNLOAD
    '========================================================

    Public Function GetUnloadOption(
        ByVal dummy As String) As Integer

        Return Session.LibraryUnloadOption.Immediately

    End Function

End Module
