Option Strict On

Imports System
Imports System.IO
Imports System.Globalization
Imports System.Collections.Generic

Imports NXOpen
Imports NXOpen.Assemblies
Imports NXOpen.Features


'==============================================================================
' NX 2306
' CUSTOMER EXTERNAL SURFACE SKIN CREATOR
'
' OUTPUT:
' <SourceAssemblyName>_CUSTOMER_SURFACE_SKIN.prt
'
' PROCESS:
'
'   Source Assembly
'        |
'        +-- Fully load assembly
'        |
'        +-- Simplify assembly
'        |     |
'        |     +-- Remove enclosed internal bodies
'        |     +-- Remove internal enclosed voids
'        |     +-- Optional small body removal
'        |     +-- Optional small hole removal
'        |     +-- Optional small blend removal
'        |
'        +-- Create simplified solid geometry
'        |
'        +-- Extract all remaining faces
'        |
'        +-- Create independent surface geometry
'        |
'        +-- Delete temporary simplified solids
'        |
'        +-- Check final solid / sheet body count
'        |
'        +-- Save customer surface model
'
'
' IMPORTANT:
'
' NX SimplifyBuilder's "internal body" logic primarily means geometry
' completely enclosed by other geometry.
'
' This is NOT a true optical / ray-tracing visibility algorithm.
'
' For example:
'
'   An internal shaft visible through a large housing opening may remain.
'
'
' SOURCE SAFETY:
'
' The original assembly is NOT modified.
'
'
' TEAMCENTER:
'
' This version creates a native/local NX .prt file.
'
' For Teamcenter Managed NX:
'
' Replace CreateDestinationPart() with the organization's approved
' Teamcenter Item / Item Revision creation process.
'
'==============================================================================


Module NX2306_CustomerExternalSurfaceSkin


    '==========================================================================
    ' NX SESSION
    '==========================================================================

    Private theSession As Session

    Private sourcePart As Part

    Private lw As ListingWindow



    '==========================================================================
    ' USER SETTINGS
    '==========================================================================


    '----------------------------------------------------------------------
    ' INTERNAL GEOMETRY
    '----------------------------------------------------------------------

    Private Const REMOVE_INTERNAL_BODIES As Boolean = True

    Private Const REMOVE_INTERNAL_VOIDS As Boolean = True



    '----------------------------------------------------------------------
    ' UNITE RESULTING BODIES
    '
    ' True:
    ' Try to unite overlapping simplified bodies.
    '
    ' False:
    ' Preserve separate bodies where possible.
    '----------------------------------------------------------------------

    Private Const UNITE_BODIES As Boolean = True



    '----------------------------------------------------------------------
    ' SMALL BODY REMOVAL
    '
    ' Conservative default = False.
    '
    ' Enable after verifying that small customer interface components
    ' are not required.
    '----------------------------------------------------------------------

    Private Const REMOVE_SMALL_BODIES As Boolean = False

    Private Const MINIMUM_BODY_SIZE_MM As Double = 3.0



    '----------------------------------------------------------------------
    ' SMALL HOLE REMOVAL
    '
    ' Conservative default = False.
    '
    ' Keep FALSE when small mounting / locating holes are important.
    '----------------------------------------------------------------------

    Private Const REMOVE_SMALL_HOLES As Boolean = False

    Private Const MAXIMUM_HOLE_DIAMETER_MM As Double = 5.0



    '----------------------------------------------------------------------
    ' SMALL BLEND REMOVAL
    '----------------------------------------------------------------------

    Private Const REMOVE_SMALL_BLENDS As Boolean = False

    Private Const MAXIMUM_BLEND_RADIUS_MM As Double = 2.0



    '----------------------------------------------------------------------
    ' EXTRACT FACE OPTION
    '
    ' False:
    ' Keep visible holes in the extracted surface.
    '
    ' True:
    ' Attempt to remove holes while extracting faces.
    '----------------------------------------------------------------------

    Private Const DELETE_HOLES_DURING_SURFACE_EXTRACTION _
        As Boolean = False



    '----------------------------------------------------------------------
    ' DELETE TEMPORARY SIMPLIFIED SOLID
    '----------------------------------------------------------------------

    Private Const DELETE_INTERMEDIATE_SOLIDS As Boolean = True



    '----------------------------------------------------------------------
    ' CUSTOMER DATA SAFETY
    '
    ' If no sheet bodies are created or solids remain,
    ' do not save the final customer model.
    '----------------------------------------------------------------------

    Private Const BLOCK_SAVE_IF_FINAL_VALIDATION_FAILS _
        As Boolean = True



    '----------------------------------------------------------------------
    ' OUTPUT FILE
    '----------------------------------------------------------------------

    Private Const OUTPUT_SUFFIX As String =
        "_CUSTOMER_SURFACE_SKIN.prt"



    '==========================================================================
    ' STATISTICS
    '==========================================================================

    Private sourceComponentOccurrences As Integer = 0

    Private sourceSolidBodyOccurrences As Integer = 0

    Private simplifiedSolidCount As Integer = 0

    Private surfaceExtractionOperations As Integer = 0

    Private finalSolidCount As Integer = 0

    Private finalSheetCount As Integer = 0

    Private warningCount As Integer = 0

    Private errorCount As Integer = 0

    Private destinationPath As String = ""



    '==========================================================================
    ' MAIN
    '==========================================================================

    Public Sub Main()


        theSession =
            Session.GetSession()


        sourcePart =
            theSession.Parts.Work


        lw =
            theSession.ListingWindow


        lw.Open()


        PrintHeader()



        '----------------------------------------------------------------------
        ' CHECK ACTIVE PART
        '----------------------------------------------------------------------

        If sourcePart Is Nothing Then

            LogError(
                "No work part is open.")

            PrintSummary(
                "ERROR")

            Return

        End If



        '----------------------------------------------------------------------
        ' CHECK ASSEMBLY
        '----------------------------------------------------------------------

        If sourcePart.ComponentAssembly Is Nothing OrElse
           sourcePart.ComponentAssembly.RootComponent Is Nothing Then


            LogError(
                "The active work part is not an assembly.")


            PrintSummary(
                "ERROR")


            Return

        End If



        '----------------------------------------------------------------------
        ' UNDO MARK
        '----------------------------------------------------------------------

        Dim undoMark As Session.UndoMarkId =

            theSession.SetUndoMark(
                Session.MarkVisibility.Visible,
                "Create Customer External Surface Skin")



        Try


            '==================================================================
            ' STEP 1
            ' FULLY LOAD SOURCE ASSEMBLY
            '==================================================================

            LogInfo(
                "STEP 1 - Fully loading source assembly...")


            EnsureSourceAssemblyFullyLoaded()



            '==================================================================
            ' STEP 2
            ' ANALYZE SOURCE
            '==================================================================

            LogInfo("")

            LogInfo(
                "STEP 2 - Analyzing source assembly...")


            AnalyzeSourceAssembly()



            destinationPath =
                BuildDestinationPath()



            LogInfo(
                "Source assembly            : " &
                sourcePart.Leaf)


            LogInfo(
                "Component occurrences      : " &
                sourceComponentOccurrences.ToString())


            LogInfo(
                "Solid body occurrences     : " &
                sourceSolidBodyOccurrences.ToString())


            LogInfo(
                "Destination                : " &
                destinationPath)



            '==================================================================
            ' CHECK OUTPUT FILE
            '==================================================================

            If File.Exists(
                destinationPath) Then


                Throw New Exception(
                    "Destination file already exists:" &
                    Environment.NewLine &
                    destinationPath)

            End If



            '==================================================================
            ' STEP 3
            ' CREATE DESTINATION PART
            '==================================================================

            LogInfo("")

            LogInfo(
                "STEP 3 - Creating destination part...")


            Dim destinationPart As Part =

                CreateDestinationPart(
                    destinationPath)



            If destinationPart Is Nothing Then

                Throw New Exception(
                    "Unable to create/find destination part.")

            End If



            EnsurePartFullyLoaded(
                destinationPart)



            '==================================================================
            ' STEP 4
            ' SIMPLIFY ASSEMBLY
            '==================================================================

            LogInfo("")

            LogInfo(
                "STEP 4 - Simplifying assembly...")


            CreateSimplifiedExternalGeometry(
                destinationPart)



            '==================================================================
            ' GET SIMPLIFIED SOLIDS
            '==================================================================

            Dim simplifiedSolids As List(Of Body) =

                GetSolidBodies(
                    destinationPart)



            simplifiedSolidCount =
                simplifiedSolids.Count



            LogInfo(
                "Simplified solid bodies    : " &
                simplifiedSolidCount.ToString())



            If simplifiedSolidCount = 0 Then

                Throw New Exception(
                    "Simplification completed but no destination " &
                    "solid bodies were found.")

            End If



            '==================================================================
            ' BODY REDUCTION INFORMATION
            '==================================================================

            Dim bodyCountReduction As Integer =

                Math.Max(
                    0,
                    sourceSolidBodyOccurrences -
                    simplifiedSolidCount)



            LogInfo(
                "Body-count reduction       : " &
                bodyCountReduction.ToString())



            LogWarning(
                "Body-count reduction is not an exact internal-body " &
                "removal count because Unite Bodies can also reduce " &
                "the number of bodies.")



            '==================================================================
            ' STEP 5
            ' EXTRACT SURFACE SKIN
            '==================================================================

            LogInfo("")

            LogInfo(
                "STEP 5 - Extracting independent surface skin...")



            Dim sheetCountBefore As Integer =

                CountSheetBodies(
                    destinationPart)



            surfaceExtractionOperations =

                ExtractSurfaceSkin(
                    destinationPart,
                    simplifiedSolids)



            Dim sheetCountAfter As Integer =

                CountSheetBodies(
                    destinationPart)



            LogInfo(
                "Extraction operations      : " &
                surfaceExtractionOperations.ToString())


            LogInfo(
                "Sheet bodies before        : " &
                sheetCountBefore.ToString())


            LogInfo(
                "Sheet bodies after         : " &
                sheetCountAfter.ToString())



            '------------------------------------------------------------------
            ' CHECK EXTRACTION
            '------------------------------------------------------------------

            If surfaceExtractionOperations = 0 Then

                Throw New Exception(
                    "No Extract Face operation completed successfully.")

            End If



            If sheetCountAfter <= sheetCountBefore Then

                Throw New Exception(
                    "Extract Face completed but no new sheet bodies " &
                    "were detected. Intermediate solids will not be deleted.")

            End If



            '==================================================================
            ' STEP 6
            ' DELETE SIMPLIFIED SOLID BODIES
            '==================================================================

            If DELETE_INTERMEDIATE_SOLIDS Then


                LogInfo("")

                LogInfo(
                    "STEP 6 - Deleting intermediate simplified solids...")


                DeleteBodiesSafely(
                    simplifiedSolids,
                    undoMark)


            Else


                LogWarning(
                    "DELETE_INTERMEDIATE_SOLIDS=False. " &
                    "Destination can contain solid bodies.")


            End If



            '==================================================================
            ' UPDATE
            '==================================================================

            Try


                Dim updateErrors As Integer =

                    theSession.UpdateManager.DoUpdate(
                        undoMark)



                If updateErrors > 0 Then


                    LogWarning(
                        "NX final update reported " &
                        updateErrors.ToString() &
                        " error(s).")


                End If


            Catch ex As Exception


                LogWarning(
                    "Final update exception: " &
                    ex.Message)


            End Try



            '==================================================================
            ' STEP 7
            ' FINAL VALIDATION
            '==================================================================

            LogInfo("")

            LogInfo(
                "STEP 7 - Validating final output...")


            CountBodyTypes(
                destinationPart,
                finalSolidCount,
                finalSheetCount)



            LogInfo(
                "Final solid bodies         : " &
                finalSolidCount.ToString())


            LogInfo(
                "Final sheet bodies         : " &
                finalSheetCount.ToString())



            Dim validationPassed As Boolean =

                ValidateFinalOutput()



            '==================================================================
            ' BLOCK UNSAFE CUSTOMER FILE
            '==================================================================

            If Not validationPassed AndAlso
               BLOCK_SAVE_IF_FINAL_VALIDATION_FAILS Then


                LogError(
                    "Final validation failed. Output was NOT saved " &
                    "because BLOCK_SAVE_IF_FINAL_VALIDATION_FAILS=True.")


                PrintSummary(
                    "ERROR")


                Return

            End If



            '==================================================================
            ' STEP 8
            ' SAVE
            '==================================================================

            LogInfo("")

            LogInfo(
                "STEP 8 - Saving destination part...")


            SaveDestinationPart(
                destinationPart)



            '==================================================================
            ' RESULT
            '==================================================================

            Dim result As String =
                "PASS"



            If errorCount > 0 Then

                result =
                    "ERROR"


            ElseIf warningCount > 0 Then

                result =
                    "PASS WITH WARNINGS"

            End If



            PrintSummary(
                result)



        Catch ex As Exception


            LogError(
                ex.Message)



            If Not String.IsNullOrWhiteSpace(
                ex.StackTrace) Then


                lw.WriteLine("")

                lw.WriteLine(
                    ex.StackTrace)


            End If



            PrintSummary(
                "ERROR")


        End Try


    End Sub



    '==========================================================================
    ' FULLY LOAD SOURCE ASSEMBLY
    '==========================================================================

    Private Sub EnsureSourceAssemblyFullyLoaded()


        Dim parts(0) As BasePart

        parts(0) =
            sourcePart



        Dim loadStatus As PartLoadStatus =
            Nothing



        Try


            loadStatus =

                theSession.Parts.EnsurePartsLoadedFully(
                    parts,
                    True)



            LogInfo(
                "Assembly load              : Completed")



        Finally


            If loadStatus IsNot Nothing Then

                loadStatus.Dispose()

            End If


        End Try


    End Sub



    '==========================================================================
    ' FULLY LOAD DESTINATION
    '==========================================================================

    Private Sub EnsurePartFullyLoaded(
        ByVal part As Part)


        Dim parts(0) As BasePart

        parts(0) =
            part



        Dim loadStatus As PartLoadStatus =
            Nothing



        Try


            loadStatus =

                theSession.Parts.EnsurePartsLoadedFully(
                    parts,
                    False)



        Finally


            If loadStatus IsNot Nothing Then

                loadStatus.Dispose()

            End If


        End Try


    End Sub



    '==========================================================================
    ' ANALYZE SOURCE ASSEMBLY
    '==========================================================================

    Private Sub AnalyzeSourceAssembly()


        sourceComponentOccurrences =
            0


        sourceSolidBodyOccurrences =
            0



        Dim root As Component =

            sourcePart.ComponentAssembly.RootComponent



        For Each child As Component _
            In root.GetChildren()


            AnalyzeComponentRecursive(
                child)


        Next


    End Sub



    '==========================================================================
    ' RECURSIVE COMPONENT ANALYSIS
    '==========================================================================

    Private Sub AnalyzeComponentRecursive(
        ByVal component As Component)


        If component Is Nothing Then

            Return

        End If



        sourceComponentOccurrences +=
            1



        '----------------------------------------------------------------------
        ' COUNT SOLID BODIES
        '----------------------------------------------------------------------

        Try


            Dim prototypePart As Part =

                TryCast(
                    component.Prototype,
                    Part)



            If prototypePart IsNot Nothing Then


                For Each body As Body _
                    In prototypePart.Bodies


                    Try


                        If body.IsSolidBody Then

                            sourceSolidBodyOccurrences +=
                                1

                        End If


                    Catch

                    End Try


                Next


            End If



        Catch ex As Exception


            LogWarning(
                "Unable to count bodies for component '" &
                SafeComponentName(component) &
                "': " &
                ex.Message)


        End Try



        '----------------------------------------------------------------------
        ' CHILD COMPONENTS
        '----------------------------------------------------------------------

        Try


            For Each child As Component _
                In component.GetChildren()


                AnalyzeComponentRecursive(
                    child)


            Next



        Catch ex As Exception


            LogWarning(
                "Unable to traverse component '" &
                SafeComponentName(component) &
                "': " &
                ex.Message)


        End Try


    End Sub



    '==========================================================================
    ' BUILD DESTINATION PATH
    '==========================================================================

    Private Function BuildDestinationPath() As String


        Dim outputFolder As String =
            ""



        Try


            If Not String.IsNullOrWhiteSpace(
                sourcePart.FullPath) Then


                outputFolder =

                    Path.GetDirectoryName(
                        sourcePart.FullPath)


            End If


        Catch

        End Try



        '----------------------------------------------------------------------
        ' FALLBACK TO DESKTOP
        '----------------------------------------------------------------------

        If String.IsNullOrWhiteSpace(
            outputFolder) OrElse
           Not Directory.Exists(
            outputFolder) Then


            outputFolder =

                Environment.GetFolderPath(
                    Environment.SpecialFolder.DesktopDirectory)


        End If



        Dim sourceBaseName As String =

            Path.GetFileNameWithoutExtension(
                sourcePart.Leaf)



        Return Path.Combine(
            outputFolder,
            sourceBaseName &
            OUTPUT_SUFFIX)


    End Function



    '==========================================================================
    ' CREATE DESTINATION PART
    '==========================================================================

    Private Function CreateDestinationPart(
        ByVal outputPath As String) As Part


        Dim fileNew As FileNew =

            theSession.Parts.FileNew()



        Try


            fileNew.UseBlankTemplate =
                True


            fileNew.TemplateType =
                FileNewTemplateType.Item


            fileNew.ApplicationName =
                "Modeling"


            fileNew.UsesMasterModel =
                "No"



            '------------------------------------------------------------------
            ' SAME UNIT SYSTEM AS SOURCE
            '------------------------------------------------------------------

            If sourcePart.PartUnits =
               BasePart.Units.Inches Then


                fileNew.Units =
                    Part.Units.Inches


            Else


                fileNew.Units =
                    Part.Units.Millimeters


            End If



            fileNew.NewFileName =
                outputPath


            fileNew.MasterFileName =
                ""


            'Keep original assembly displayed.
            fileNew.MakeDisplayedPart =
                False



            fileNew.Commit()



        Finally


            fileNew.Destroy()


        End Try



        '----------------------------------------------------------------------
        ' FIND NEW PART IN SESSION
        '----------------------------------------------------------------------

        For Each basePart As BasePart _
            In theSession.Parts


            Dim candidate As Part =

                TryCast(
                    basePart,
                    Part)



            If candidate Is Nothing Then

                Continue For

            End If



            '------------------------------------------------------------------
            ' FULL PATH MATCH
            '----------------------------------------------------------------------

            Try


                If String.Equals(
                    candidate.FullPath,
                    outputPath,
                    StringComparison.OrdinalIgnoreCase) Then


                    Return candidate


                End If


            Catch

            End Try



            '------------------------------------------------------------------
            ' LEAF NAME FALLBACK
            '----------------------------------------------------------------------

            Try


                If String.Equals(
                    candidate.Leaf,
                    Path.GetFileName(outputPath),
                    StringComparison.OrdinalIgnoreCase) Then


                    Return candidate


                End If


            Catch

            End Try


        Next



        Return Nothing


    End Function



    '==========================================================================
    ' SIMPLIFY ASSEMBLY
    '==========================================================================

    Private Sub CreateSimplifiedExternalGeometry(
        ByVal destinationPart As Part)


        Dim builder As SimplifyBuilder =
            Nothing



        Try


            builder =

                sourcePart.AssemblyManager.
                    CreateSimplifyBuilder()



            '------------------------------------------------------------------
            ' DESTINATION PART
            '----------------------------------------------------------------------

            builder.SetDestinationPart(
                destinationPart)



            '------------------------------------------------------------------
            ' SOURCE COMPONENTS
            '----------------------------------------------------------------------

            Dim root As Component =

                sourcePart.ComponentAssembly.RootComponent



            Dim selectedCount As Integer =
                0



            For Each child As Component _
                In root.GetChildren()


                If child Is Nothing Then

                    Continue For

                End If



                Try


                    If builder.ObjectsToSimplify.Add(
                        child) Then


                        selectedCount +=
                            1


                    End If



                Catch ex As Exception


                    LogWarning(
                        "Could not add component '" &
                        SafeComponentName(child) &
                        "' to SimplifyBuilder: " &
                        ex.Message)


                End Try


            Next



            If selectedCount = 0 Then


                Throw New Exception(
                    "No source components were added to SimplifyBuilder.")


            End If



            LogInfo(
                "Components sent to Simplify: " &
                selectedCount.ToString())



            '==================================================================
            ' CORE EXTERNAL GEOMETRY RULES
            '==================================================================

            builder.RemoveInternalBodies =
                REMOVE_INTERNAL_BODIES


            builder.RemoveInternalVoids =
                REMOVE_INTERNAL_VOIDS


            builder.UniteBodies =
                UNITE_BODIES



            '==================================================================
            ' SMALL BODY FILTER
            '==================================================================

            builder.UseBodySize =
                REMOVE_SMALL_BODIES



            If REMOVE_SMALL_BODIES Then


                builder.MinimumBodySize.RightHandSide =

                    ToExpressionValue(
                        ConvertMmToSourcePartUnits(
                            MINIMUM_BODY_SIZE_MM))


            End If



            '==================================================================
            ' SMALL HOLE FILTER
            '==================================================================

            builder.UseHoleDiameter =
                REMOVE_SMALL_HOLES



            If REMOVE_SMALL_HOLES Then


                builder.MaximumHoleSize.RightHandSide =

                    ToExpressionValue(
                        ConvertMmToSourcePartUnits(
                            MAXIMUM_HOLE_DIAMETER_MM))


            End If



            '==================================================================
            ' SMALL BLEND FILTER
            '==================================================================

            builder.UseBlendRadius =
                REMOVE_SMALL_BLENDS



            If REMOVE_SMALL_BLENDS Then


                builder.MaximumBlendRadius.RightHandSide =

                    ToExpressionValue(
                        ConvertMmToSourcePartUnits(
                            MAXIMUM_BLEND_RADIUS_MM))


            End If



            '==================================================================
            ' VALIDATE
            '==================================================================

            If Not builder.Validate() Then


                Throw New Exception(
                    "SimplifyBuilder.Validate() returned False.")


            End If



            '==================================================================
            ' COMMIT
            '==================================================================

            builder.Commit()



            LogInfo(
                "SimplifyBuilder             : Commit completed")



        Finally


            If builder IsNot Nothing Then


                Try

                    builder.Destroy()

                Catch

                End Try


            End If


        End Try


    End Sub



    '==========================================================================
    ' GET SOLID BODIES
    '==========================================================================

    Private Function GetSolidBodies(
        ByVal part As Part) As List(Of Body)


        Dim result As New List(Of Body)



        For Each body As Body _
            In part.Bodies


            Try


                If body.IsSolidBody Then

                    result.Add(
                        body)

                End If



            Catch ex As Exception


                LogWarning(
                    "Could not classify Body Tag " &
                    body.Tag.ToString() &
                    ": " &
                    ex.Message)


            End Try


        Next



        Return result


    End Function



    '==========================================================================
    ' EXTRACT EXTERNAL SURFACE SKIN
    '==========================================================================

    Private Function ExtractSurfaceSkin(
        ByVal destinationPart As Part,
        ByVal sourceSolids As List(Of Body)) As Integer


        Dim successfulOperations As Integer =
            0



        For Each solidBody As Body _
            In sourceSolids


            If solidBody Is Nothing Then

                Continue For

            End If



            Dim builder As ExtractFaceBuilder =
                Nothing



            Try


                builder =

                    destinationPart.Features.
                        CreateExtractFaceBuilder(
                            CType(
                                Nothing,
                                Feature))



                '==================================================================
                ' EXTRACT FACES
                '==================================================================

                builder.Type =

                    ExtractFaceBuilder.
                        ExtractType.Face



                builder.FaceOption =

                    ExtractFaceBuilder.
                        FaceOptionType.AllBodyFaces



                builder.SurfaceType =

                    ExtractFaceBuilder.
                        FaceSurfaceType.SameAsOriginal



                '------------------------------------------------------------------
                ' INDEPENDENT RESULT
                '
                ' Required because source solid will be deleted.
                '----------------------------------------------------------------------

                builder.Associative =
                    False



                builder.HideOriginal =
                    False



                builder.InheritDisplayProperties =
                    True



                builder.InheritMaterial =
                    False



                builder.CopyThreads =
                    False



                builder.DeleteHoles =

                    DELETE_HOLES_DURING_SURFACE_EXTRACTION



                builder.FeatureOption =

                    ExtractFaceBuilder.
                        FeatureOptionType.
                        OneFeatureForAllBodies



                '==================================================================
                ' ADD SOLID BODY
                '==================================================================

                If Not builder.ObjectToExtract.Add(
                    CType(
                        solidBody,
                        DisplayableObject)) Then


                    LogWarning(
                        "Body Tag " &
                        solidBody.Tag.ToString() &
                        " was not added to ExtractFaceBuilder.")


                    Continue For


                End If



                '==================================================================
                ' VALIDATE
                '==================================================================

                If Not builder.Validate() Then


                    LogWarning(
                        "ExtractFaceBuilder.Validate() failed for Body Tag " &
                        solidBody.Tag.ToString())


                    Continue For


                End If



                '==================================================================
                ' COMMIT
                '==================================================================

                Dim extractedFeature As Feature =

                    builder.CommitFeature()



                If extractedFeature IsNot Nothing Then


                    successfulOperations +=
                        1



                    LogInfo(
                        "Extracted surface from Body Tag " &
                        solidBody.Tag.ToString())


                Else


                    LogWarning(
                        "Extract Face returned no feature for Body Tag " &
                        solidBody.Tag.ToString())


                End If



            Catch ex As Exception


                LogWarning(
                    "Surface extraction failed for Body Tag " &
                    solidBody.Tag.ToString() &
                    ": " &
                    ex.Message)



            Finally


                If builder IsNot Nothing Then


                    Try

                        builder.Destroy()

                    Catch

                    End Try


                End If


            End Try


        Next



        Return successfulOperations


    End Function



    '==========================================================================
    ' DELETE INTERMEDIATE SOLID BODIES
    '==========================================================================

    Private Sub DeleteBodiesSafely(
        ByVal bodies As List(Of Body),
        ByVal undoMark As Session.UndoMarkId)


        If bodies Is Nothing OrElse
           bodies.Count = 0 Then


            Return


        End If



        Dim deleteList As New List(Of TaggedObject)



        For Each body As Body _
            In bodies


            If body Is Nothing Then

                Continue For

            End If



            Try


                If body.IsSolidBody Then


                    deleteList.Add(
                        body)


                End If


            Catch

            End Try


        Next



        If deleteList.Count = 0 Then

            Return

        End If



        theSession.UpdateManager.
            ClearErrorList()



        '----------------------------------------------------------------------
        ' CURRENT NX API
        '----------------------------------------------------------------------

        Dim addErrors As Integer =

            theSession.UpdateManager.
                AddObjectsToDeleteList(
                    deleteList.ToArray())



        If addErrors > 0 Then


            LogWarning(
                "AddObjectsToDeleteList reported " &
                addErrors.ToString() &
                " error(s).")


        End If



        '----------------------------------------------------------------------
        ' UPDATE
        '----------------------------------------------------------------------

        Dim updateErrors As Integer =

            theSession.UpdateManager.
                DoUpdate(
                    undoMark)



        If updateErrors > 0 Then


            LogWarning(
                "Deleting temporary solids produced " &
                updateErrors.ToString() &
                " update error(s).")


        Else


            LogInfo(
                "Temporary solids deleted   : " &
                deleteList.Count.ToString())


        End If


    End Sub



    '==========================================================================
    ' COUNT SHEET BODIES
    '==========================================================================

    Private Function CountSheetBodies(
        ByVal part As Part) As Integer


        Dim count As Integer =
            0



        For Each body As Body _
            In part.Bodies


            Try


                If body.IsSheetBody Then

                    count +=
                        1

                End If


            Catch

            End Try


        Next



        Return count


    End Function



    '==========================================================================
    ' COUNT BODY TYPES
    '==========================================================================

    Private Sub CountBodyTypes(
        ByVal part As Part,
        ByRef solidCount As Integer,
        ByRef sheetCount As Integer)


        solidCount =
            0


        sheetCount =
            0



        For Each body As Body _
            In part.Bodies


            Try


                If body.IsSolidBody Then


                    solidCount +=
                        1


                ElseIf body.IsSheetBody Then


                    sheetCount +=
                        1


                End If



            Catch ex As Exception


                LogWarning(
                    "Unable to classify final Body Tag " &
                    body.Tag.ToString() &
                    ": " &
                    ex.Message)


            End Try


        Next


    End Sub



    '==========================================================================
    ' FINAL VALIDATION
    '==========================================================================

    Private Function ValidateFinalOutput() As Boolean


        Dim passed As Boolean =
            True



        '----------------------------------------------------------------------
        ' MUST HAVE SHEET BODY
        '----------------------------------------------------------------------

        If finalSheetCount <= 0 Then


            LogError(
                "Final validation: no sheet bodies exist.")


            passed =
                False


        Else


            LogInfo(
                "PASS - Surface/sheet geometry exists.")


        End If



        '----------------------------------------------------------------------
        ' SOLIDS SHOULD BE REMOVED
        '----------------------------------------------------------------------

        If DELETE_INTERMEDIATE_SOLIDS AndAlso
           finalSolidCount > 0 Then


            LogError(
                "Final validation: " &
                finalSolidCount.ToString() &
                " solid body/bodies remain.")


            passed =
                False



        ElseIf finalSolidCount = 0 Then


            LogInfo(
                "PASS - No final solid bodies remain.")


        End If



        Return passed


    End Function



    '==========================================================================
    ' SAVE DESTINATION
    '==========================================================================

    Private Sub SaveDestinationPart(
        ByVal destinationPart As Part)


        Dim saveStatus As PartSaveStatus =
            Nothing



        Try


            saveStatus =

                destinationPart.Save(
                    BasePart.SaveComponents.False,
                    BasePart.CloseAfterSave.False)



            LogInfo(
                "Saved destination          : " &
                destinationPath)



        Finally


            If saveStatus IsNot Nothing Then

                saveStatus.Dispose()

            End If


        End Try


    End Sub



    '==========================================================================
    ' UNIT CONVERSION
    '==========================================================================

    Private Function ConvertMmToSourcePartUnits(
        ByVal valueMm As Double) As Double


        If sourcePart.PartUnits =
           BasePart.Units.Inches Then


            Return valueMm /
                   25.4


        End If



        Return valueMm


    End Function



    '==========================================================================
    ' EXPRESSION VALUE
    '==========================================================================

    Private Function ToExpressionValue(
        ByVal value As Double) As String


        Return value.ToString(
            "0.############",
            CultureInfo.InvariantCulture)


    End Function



    '==========================================================================
    ' SAFE COMPONENT NAME
    '==========================================================================

    Private Function SafeComponentName(
        ByVal component As Component) As String


        If component Is Nothing Then

            Return "<Nothing>"

        End If



        Try


            If Not String.IsNullOrWhiteSpace(
                component.DisplayName) Then


                Return component.DisplayName


            End If


        Catch

        End Try



        Try

            Return component.Name

        Catch

        End Try



        Return "Tag " &
               component.Tag.ToString()


    End Function



    '==========================================================================
    ' HEADER
    '==========================================================================

    Private Sub PrintHeader()


        lw.WriteLine("")

        lw.WriteLine(
            "==============================================================")


        lw.WriteLine(
            "  NX 2306 - CUSTOMER EXTERNAL SURFACE SKIN CREATOR")


        lw.WriteLine(
            "==============================================================")


        lw.WriteLine("")



        lw.WriteLine(
            "Settings:")


        lw.WriteLine(
            "  Remove internal bodies : " &
            REMOVE_INTERNAL_BODIES.ToString())


        lw.WriteLine(
            "  Remove internal voids  : " &
            REMOVE_INTERNAL_VOIDS.ToString())


        lw.WriteLine(
            "  Unite bodies           : " &
            UNITE_BODIES.ToString())


        lw.WriteLine(
            "  Remove small bodies    : " &
            REMOVE_SMALL_BODIES.ToString())


        lw.WriteLine(
            "  Remove small holes     : " &
            REMOVE_SMALL_HOLES.ToString())


        lw.WriteLine(
            "  Remove small blends    : " &
            REMOVE_SMALL_BLENDS.ToString())


        lw.WriteLine("")


    End Sub



    '==========================================================================
    ' LOG INFO
    '==========================================================================

    Private Sub LogInfo(
        ByVal message As String)


        lw.WriteLine(
            message)


    End Sub



    '==========================================================================
    ' LOG WARNING
    '==========================================================================

    Private Sub LogWarning(
        ByVal message As String)


        warningCount +=
            1


        lw.WriteLine(
            "[WARNING] " &
            message)


    End Sub



    '==========================================================================
    ' LOG ERROR
    '==========================================================================

    Private Sub LogError(
        ByVal message As String)


        errorCount +=
            1


        lw.WriteLine(
            "[ERROR] " &
            message)


    End Sub



    '==========================================================================
    ' SUMMARY
    '==========================================================================

    Private Sub PrintSummary(
        ByVal status As String)


        lw.WriteLine("")

        lw.WriteLine(
            "==============================================================")


        lw.WriteLine(
            "                         SUMMARY")


        lw.WriteLine(
            "==============================================================")



        If sourcePart IsNot Nothing Then


            lw.WriteLine(
                "Source assembly            : " &
                sourcePart.Leaf)


        End If



        lw.WriteLine(
            "Source component occurrences: " &
            sourceComponentOccurrences.ToString())


        lw.WriteLine(
            "Source solid occurrences   : " &
            sourceSolidBodyOccurrences.ToString())


        lw.WriteLine(
            "Simplified solids created  : " &
            simplifiedSolidCount.ToString())


        lw.WriteLine(
            "Surface extraction ops     : " &
            surfaceExtractionOperations.ToString())


        lw.WriteLine(
            "Final solid bodies         : " &
            finalSolidCount.ToString())


        lw.WriteLine(
            "Final sheet bodies         : " &
            finalSheetCount.ToString())


        lw.WriteLine(
            "Warnings                   : " &
            warningCount.ToString())


        lw.WriteLine(
            "Errors                     : " &
            errorCount.ToString())



        If Not String.IsNullOrWhiteSpace(
            destinationPath) Then


            lw.WriteLine(
                "Destination                : " &
                destinationPath)


        End If



        lw.WriteLine(
            "--------------------------------------------------------------")


        lw.WriteLine(
            "RESULT                      : " &
            status)


        lw.WriteLine(
            "==============================================================")


        lw.WriteLine("")


    End Sub



    '==========================================================================
    ' UNLOAD OPTION
    '==========================================================================

    Public Function GetUnloadOption(
        ByVal dummy As String) As Integer


        Return Session.LibraryUnloadOption.Immediately


    End Function


End Module
