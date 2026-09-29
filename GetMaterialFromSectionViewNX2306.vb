Imports NXOpen
Imports NXOpen.Assemblies
Imports NXOpen.Drawings
Imports NXOpen.UF

Module GetMaterialFromSectionViewNX2306
    
    Dim theSession As Session = Session.GetSession()
    Dim theUfSession As UFSession = UFSession.GetUFSession()
    Dim lw As ListingWindow = theSession.ListingWindow
    
    Sub Main()
        
        Try
            lw.Open()
            lw.WriteLine("===== NX 2306 - Get Material from Section View =====")
            lw.WriteLine("")
            
            ' Get the active part
            Dim workPart As Part = theSession.Parts.Work
            If workPart Is Nothing Then
                lw.WriteLine("ERROR: No active part")
                Return
            End If
            
            ' Check if it's a drawing
            Dim drawingPart As DrawingSheet = TryCast(workPart, DrawingSheet)
            If drawingPart Is Nothing Then
                lw.WriteLine("ERROR: Active part is not a drawing")
                Return
            End If
            
            lw.WriteLine("Drawing Part: " & drawingPart.Name)
            lw.WriteLine("")
            
            ' Get all section views in the drawing
            Dim sectionViews As New List(Of DrawingSheetView)
            
            For Each view In drawingPart.DrawingViews
                If view.ViewType = DrawingSheetView.ViewTypes.Section Then
                    sectionViews.Add(view)
                End If
            Next
            
            If sectionViews.Count = 0 Then
                lw.WriteLine("No section views found in this drawing")
                Return
            End If
            
            lw.WriteLine("Section Views Found: " & sectionViews.Count)
            lw.WriteLine("")
            
            ' Process each section view
            For Each sectionView In sectionViews
                ProcessSectionView(sectionView)
            Next
            
            lw.WriteLine("===== Process Complete =====")
            
        Catch ex As Exception
            lw.WriteLine("ERROR: " & ex.Message)
            lw.WriteLine(ex.StackTrace)
        End Try
        
    End Sub
    
    ' Process individual section view
    Sub ProcessSectionView(sectionView As DrawingSheetView)
        
        Try
            lw.WriteLine("========================================")
            lw.WriteLine("Section View: " & sectionView.Name)
            lw.WriteLine("Label: " & sectionView.Label)
            lw.WriteLine("========================================")
            
            ' Get the source view (the main view that this section was cut from)
            Dim sourceView As DrawingSheetView = sectionView.SourceView
            If sourceView Is Nothing Then
                lw.WriteLine("No source view found")
                Return
            End If
            
            lw.WriteLine("Source View: " & sourceView.Name)
            
            ' Get the origin model (assembly)
            Dim originModel As Part = sourceView.OriginModel
            If originModel Is Nothing Then
                lw.WriteLine("No origin model found")
                Return
            End If
            
            lw.WriteLine("Origin Model: " & originModel.Name)
            lw.WriteLine("")
            
            ' Get all visible components in the section
            Dim visibleComponents As List(Of Component) = GetVisibleComponentsInSection(sectionView, originModel)
            
            If visibleComponents.Count = 0 Then
                lw.WriteLine("No components visible in this section")
                Return
            End If
            
            lw.WriteLine("Visible Components: " & visibleComponents.Count)
            lw.WriteLine("")
            
            ' Extract and display material for each component
            For Each component In visibleComponents
                ExtractComponentMaterial(component)
            Next
            
            lw.WriteLine("")
            
        Catch ex As Exception
            lw.WriteLine("ERROR in ProcessSectionView: " & ex.Message)
        End Try
        
    End Sub
    
    ' Get visible components in the section view
    Function GetVisibleComponentsInSection(sectionView As DrawingSheetView, assembly As Part) As List(Of Component)
        
        Dim visibleComponents As New List(Of Component)
        
        Try
            ' Cast to component assembly
            If Not TypeOf assembly Is ComponentAssembly Then
                Return visibleComponents
            End If
            
            Dim compAssembly As ComponentAssembly = CType(assembly, ComponentAssembly)
            Dim rootComponent As Component = compAssembly.RootComponent
            
            ' Get all components from the assembly
            Dim allComponents As New List(Of Component)
            CollectAllComponents(rootComponent, allComponents)
            
            ' Filter for visible components (optional - you can refine this)
            For Each comp In allComponents
                If comp.Visible Then
                    visibleComponents.Add(comp)
                End If
            Next
            
            ' If no visible components, return all components
            If visibleComponents.Count = 0 Then
                visibleComponents = allComponents
            End If
            
            Return visibleComponents
            
        Catch ex As Exception
            lw.WriteLine("ERROR in GetVisibleComponentsInSection: " & ex.Message)
            Return visibleComponents
        End Try
        
    End Function
    
    ' Recursively collect all components from the assembly tree
    Sub CollectAllComponents(component As Component, componentList As List(Of Component))
        
        Try
            If component Is Nothing Then
                Return
            End If
            
            ' Add component if it's not the root
            Try
                Dim prototypePart As Part = component.Prototype.OwningPart
                If prototypePart IsNot Nothing AndAlso Not IsRootComponent(component) Then
                    componentList.Add(component)
                End If
            Catch
                ' Skip if error occurs
            End Try
            
            ' Recursively add child components
            Try
                Dim children As Component() = component.GetChildren()
                If children IsNot Nothing Then
                    For Each childComp In children
                        CollectAllComponents(childComp, componentList)
                    Next
                End If
            Catch
                ' Skip if no children
            End Try
            
        Catch ex As Exception
            ' Continue processing
        End Try
        
    End Sub
    
    ' Check if component is root
    Function IsRootComponent(component As Component) As Boolean
        Try
            Return component.IsSuppressed = False AndAlso component.ComponentType = Component.ComponentTypes.Part
        Catch
            Return False
        End Try
    End Function
    
    ' Extract and display material information for a component
    Sub ExtractComponentMaterial(component As Component)
        
        Try
            Dim prototypePart As Part = component.Prototype.OwningPart
            
            lw.WriteLine("Component: " & component.Name)
            lw.WriteLine("  Type: " & component.ComponentType.ToString())
            lw.WriteLine("  Reference Set: " & component.ReferenceSet)
            lw.WriteLine("  Part Name: " & prototypePart.Name)
            
            ' Get material using multiple methods
            Dim material As String = ""
            
            ' Method 1: From part attributes (NX 2306)
            material = GetMaterialFromAttributes(prototypePart)
            If material <> "" Then
                lw.WriteLine("  Material (Attribute): " & material)
            End If
            
            ' Method 2: From custom user attributes
            Dim userMaterial As String = GetMaterialFromUserAttributes(prototypePart)
            If userMaterial <> "" Then
                lw.WriteLine("  Material (User Attr): " & userMaterial)
            End If
            
            ' Method 3: From part description
            Dim descMaterial As String = GetMaterialFromDescription(prototypePart)
            If descMaterial <> "" Then
                lw.WriteLine("  Material (Description): " & descMaterial)
            End If
            
            ' Method 4: From mass properties (if available)
            Dim massProps As String = GetMaterialFromMassProperties(prototypePart)
            If massProps <> "" Then
                lw.WriteLine("  Material (Mass Props): " & massProps)
            End If
            
            If material = "" AndAlso userMaterial = "" AndAlso descMaterial = "" Then
                lw.WriteLine("  Material: Not found")
            End If
            
            lw.WriteLine("")
            
        Catch ex As Exception
            lw.WriteLine("ERROR extracting material for " & component.Name & ": " & ex.Message)
            lw.WriteLine("")
        End Try
        
    End Sub
    
    ' Get material from standard part attributes (NX 2306)
    Function GetMaterialFromAttributes(part As Part) As String
        
        Try
            Dim attrContainer As AttributeContainer = CType(part, AttributeContainer)
            If attrContainer Is Nothing Then
                Return ""
            End If
            
            ' Common material attribute names for NX 2306
            Dim materialAttrs As String() = {
                "Material",
                "MATERIAL", 
                "material",
                "Part_Material",
                "part_material",
                "MaterialName",
                "material_name",
                "NX_MATERIAL"
            }
            
            For Each attrName In materialAttrs
                Try
                    Dim attr As NXObject = attrContainer.GetAttribute(attrName)
                    If attr IsNot Nothing Then
                        Return GetAttributeValue(attr)
                    End If
                Catch
                    ' Continue to next
                End Try
            Next
            
            Return ""
            
        Catch ex As Exception
            Return ""
        End Try
        
    End Function
    
    ' Get material from user-defined attributes
    Function GetMaterialFromUserAttributes(part As Part) As String
        
        Try
            ' Get all attributes from the part
            Dim attrContainer As AttributeContainer = CType(part, AttributeContainer)
            If attrContainer Is Nothing Then
                Return ""
            End If
            
            ' This works in NX 2306 to get user attributes
            Try
                Dim materialAttr As NXObject = attrContainer.GetAttribute("USER_MATERIAL")
                If materialAttr IsNot Nothing Then
                    Return GetAttributeValue(materialAttr)
                End If
            Catch
            End Try
            
            Return ""
            
        Catch ex As Exception
            Return ""
        End Try
        
    End Function
    
    ' Get material from part description/title
    Function GetMaterialFromDescription(part As Part) As String
        
        Try
            ' Check part description
            If part.DescriptorString <> "" Then
                Dim desc As String = part.DescriptorString
                If desc.Contains("Material") Or desc.Contains("material") Then
                    Return desc
                End If
            End If
            
            ' Check name for material hints
            If part.Name.Contains("_") Then
                Dim parts As String() = part.Name.Split("_"c)
                For Each p In parts
                    If p.Length > 2 AndAlso Not p.All(Function(c) Char.IsDigit(c)) Then
                        Return p
                    End If
                Next
            End If
            
            Return ""
            
        Catch ex As Exception
            Return ""
        End Try
        
    End Function
    
    ' Get material from mass properties
    Function GetMaterialFromMassProperties(part As Part) As String
        
        Try
            ' Try to access mass properties if available
            If TypeOf part Is Part Then
                Dim massProps As NXObject = TryCast(part, NXObject)
                If massProps IsNot Nothing Then
                    ' Material density or properties might be stored
                    Dim attrContainer As AttributeContainer = CType(part, AttributeContainer)
                    Try
                        Dim densityAttr As NXObject = attrContainer.GetAttribute("Density")
                        If densityAttr IsNot Nothing Then
                            Return "Material with Density: " & GetAttributeValue(densityAttr)
                        End If
                    Catch
                    End Try
                End If
            End If
            
            Return ""
            
        Catch ex As Exception
            Return ""
        End Try
        
    End Function
    
    ' Helper function to get attribute value as string
    Function GetAttributeValue(attr As NXObject) As String
        
        Try
            If attr Is Nothing Then
                Return ""
            End If
            
            Dim strAttr As StringAttribute = TryCast(attr, StringAttribute)
            If strAttr IsNot Nothing Then
                Return strAttr.StringValue
            End If
            
            Dim intAttr As IntegerAttribute = TryCast(attr, IntegerAttribute)
            If intAttr IsNot Nothing Then
                Return intAttr.IntegerValue.ToString()
            End If
            
            Dim realAttr As RealAttribute = TryCast(attr, RealAttribute)
            If realAttr IsNot Nothing Then
                Return realAttr.RealValue.ToString("F2")
            End If
            
            Return attr.ToString()
            
        Catch ex As Exception
            Return ""
        End Try
        
    End Function
    
End Module
