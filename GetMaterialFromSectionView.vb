Imports NXOpen
Imports NXOpen.Assemblies
Imports NXOpen.UF

Module GetMaterialFromSectionView
    
    Dim theSession As Session = Session.GetSession()
    Dim theUfSession As UFSession = UFSession.GetUFSession()
    Dim lw As ListingWindow = theSession.ListingWindow
    
    Sub Main()
        
        Try
            lw.Open()
            
            ' Get the active part/component
            Dim workPart As Part = theSession.Parts.Work
            If workPart Is Nothing Then
                lw.WriteLine("ERROR: No active part")
                Return
            End If
            
            ' Get all section views in the part
            Dim sectionViews As DrawingSheetView() = GetAllSectionViews(workPart)
            
            If sectionViews.Length = 0 Then
                lw.WriteLine("No section views found")
                Return
            End If
            
            lw.WriteLine("Section Views Found: " & sectionViews.Length)
            lw.WriteLine("")
            
            ' Process each section view
            For Each sectionView In sectionViews
                ProcessSectionView(sectionView)
            Next
            
        Catch ex As Exception
            lw.WriteLine("Error: " & ex.Message)
        End Try
        
    End Sub
    
    ' Get all section views from the part
    Function GetAllSectionViews(workPart As Part) As DrawingSheetView()
        
        Try
            Dim sectionViewList As New List(Of DrawingSheetView)
            
            ' Check if it's a drawing
            Dim drawingPart As DrawingSheet = TryCast(workPart, DrawingSheet)
            If drawingPart Is Nothing Then
                lw.WriteLine("Active part is not a drawing")
                Return sectionViewList.ToArray()
            End If
            
            ' Get all views from the drawing
            For Each view In drawingPart.DrawingViews
                ' Check if view is a section view
                If view.ViewType = DrawingSheetView.ViewTypes.Section Then
                    sectionViewList.Add(view)
                End If
            Next
            
            Return sectionViewList.ToArray()
            
        Catch ex As Exception
            lw.WriteLine("Error in GetAllSectionViews: " & ex.Message)
            Return New DrawingSheetView() {}
        End Try
        
    End Function
    
    ' Process individual section view to get component materials
    Sub ProcessSectionView(sectionView As DrawingSheetView)
        
        Try
            lw.WriteLine("========================================")
            lw.WriteLine("Section View: " & sectionView.Name)
            lw.WriteLine("========================================")
            
            ' Get the source view (the view that was cut for the section)
            Dim sourceView As DrawingSheetView = sectionView.SourceView
            If sourceView Is Nothing Then
                lw.WriteLine("No source view found for this section view")
                Return
            End If
            
            lw.WriteLine("Source View: " & sourceView.Name)
            
            ' Get the origin model (assembly) from the view
            Dim originModel As Part = sourceView.OriginModel
            If originModel Is Nothing Then
                lw.WriteLine("No origin model found")
                Return
            End If
            
            lw.WriteLine("Origin Model: " & originModel.Filename)
            
            ' Get all components visible in the section view
            Dim visibleComponents As Component() = GetComponentsInView(sourceView, originModel)
            
            If visibleComponents.Length = 0 Then
                lw.WriteLine("No components found in section view")
                Return
            End If
            
            lw.WriteLine("")
            lw.WriteLine("Components in Section View: " & visibleComponents.Length)
            lw.WriteLine("")
            
            ' Get material for each component
            For Each component In visibleComponents
                GetComponentMaterial(component)
            Next
            
            lw.WriteLine("")
            
        Catch ex As Exception
            lw.WriteLine("Error in ProcessSectionView: " & ex.Message)
        End Try
        
    End Sub
    
    ' Get components from the view
    Function GetComponentsInView(view As DrawingSheetView, assembly As Part) As Component()
        
        Try
            Dim componentList As New List(Of Component)
            
            ' If assembly has components, add them
            If TypeOf assembly Is Assemblies.ComponentAssembly Then
                Dim compAssembly As Assemblies.ComponentAssembly = CType(assembly, Assemblies.ComponentAssembly)
                Dim rootComp As Component = compAssembly.RootComponent
                
                ' Get all child components
                AddComponentsRecursively(rootComp, componentList)
            End If
            
            Return componentList.ToArray()
            
        Catch ex As Exception
            lw.WriteLine("Error in GetComponentsInView: " & ex.Message)
            Return New Component() {}
        End Try
        
    End Function
    
    ' Recursively add components from hierarchy
    Sub AddComponentsRecursively(component As Component, componentList As List(Of Component))
        
        Try
            If component Is Nothing Then
                Return
            End If
            
            ' Add current component if it's not the root
            If component.Prototype.OwningPart.Tag <> component.Prototype.OwningPart.OwningAssembly.RootComponent.Prototype.OwningPart.Tag Then
                componentList.Add(component)
            End If
            
            ' Recursively add children
            For Each childComponent In component.GetChildren()
                AddComponentsRecursively(childComponent, componentList)
            Next
            
        Catch ex As Exception
            ' Silently continue
        End Try
        
    End Sub
    
    ' Get material of a component
    Sub GetComponentMaterial(component As Component)
        
        Try
            ' Get the prototype part of the component
            Dim prototypePart As Part = component.Prototype.OwningPart
            
            lw.WriteLine("Component: " & component.Name)
            lw.WriteLine("  Reference Set: " & component.ReferenceSet)
            lw.WriteLine("  Component Part: " & prototypePart.Filename)
            
            ' Method 1: Get material from part properties
            Dim materialName As String = GetMaterialFromPartAttributes(prototypePart)
            If materialName <> "" Then
                lw.WriteLine("  Material (from attributes): " & materialName)
            End If
            
            ' Method 2: Get material from features
            Dim materialFromFeatures As String = GetMaterialFromFeatures(prototypePart)
            If materialFromFeatures <> "" Then
                lw.WriteLine("  Material (from features): " & materialFromFeatures)
            End If
            
            lw.WriteLine("")
            
        Catch ex As Exception
            lw.WriteLine("Error getting material for " & component.Name & ": " & ex.Message)
        End Try
        
    End Sub
    
    ' Get material from part attributes
    Function GetMaterialFromPartAttributes(part As Part) As String
        
        Try
            ' Try to get material from custom attributes
            Dim attrContainer As AttributeContainer = CType(part, AttributeContainer)
            
            ' Common material attribute names
            Dim materialAttributes As String() = {"Material", "MATERIAL", "material", "Part_Material"}
            
            For Each attrName In materialAttributes
                Try
                    Dim attr As NXObject = attrContainer.GetAttribute(attrName)
                    If attr IsNot Nothing Then
                        Dim intAttr As IntegerAttribute = TryCast(attr, IntegerAttribute)
                        Dim realAttr As RealAttribute = TryCast(attr, RealAttribute)
                        Dim strAttr As StringAttribute = TryCast(attr, StringAttribute)
                        
                        If strAttr IsNot Nothing Then
                            Return strAttr.StringValue
                        ElseIf intAttr IsNot Nothing Then
                            Return intAttr.IntegerValue.ToString()
                        ElseIf realAttr IsNot Nothing Then
                            Return realAttr.RealValue.ToString()
                        End If
                    End If
                Catch
                    ' Continue to next attribute
                End Try
            Next
            
            Return ""
            
        Catch ex As Exception
            lw.WriteLine("  Error reading attributes: " & ex.Message)
            Return ""
        End Try
        
    End Function
    
    ' Get material from part features (like material block features if they exist)
    Function GetMaterialFromFeatures(part As Part) As String
        
        Try
            ' Iterate through all features in the part
            For Each feature In part.Features
                ' Check feature type
                Dim featureName As String = feature.GetType().Name
                If featureName.Contains("Material") Or featureName.Contains("Body") Then
                    Return feature.Name
                End If
            Next
            
            Return ""
            
        Catch ex As Exception
            Return ""
        End Try
        
    End Function
    
End Module
