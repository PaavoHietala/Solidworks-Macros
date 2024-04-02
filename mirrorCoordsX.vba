Sub main()

    Dim swApp As SldWorks.SldWorks
    Set swApp = Application.SldWorks
    
    Dim swModel As SldWorks.ModelDoc2
    Set swModel = swApp.ActiveDoc
    
    Dim selectedFeatures() As SldWorks.Feature
    
    Dim swFeat As SldWorks.Feature
    
    If Not swModel Is Nothing Then
        
        Set swSelMgr = swModel.SelectionManager
        
        ' Store selected features in an array
        For i = 1 To swSelMgr.GetSelectedObjectCount2(-1)
            Set swFeat = swSelMgr.GetSelectedObject6(i, -1)
            If Not swFeat Is Nothing Then
                If swFeat.GetTypeName2() = "CoordSys" Then
                    ReDim Preserve selectedFeatures(i - 1)
                    Set selectedFeatures(i - 1) = swFeat
                Else
                    MsgBox "Selected feature is not a coordinate system"
                End If
            End If
        Next i
        
        ' Process each selected feature
        If Not IsEmpty(selectedFeatures) Then
            For i = LBound(selectedFeatures) To UBound(selectedFeatures)
                Set swFeat = selectedFeatures(i)
                
                MirrorCoordSystem swFeat, swModel
            Next i
        Else
            MsgBox "Please select coordinate system features"
        End If
        
    Else
        MsgBox "Please open model"
    End If
    
End Sub

Sub MirrorCoordSystem(selectedCoordsFeat As SldWorks.Feature, swModel As SldWorks.ModelDoc2)
    ' Get transformation matrix for the selected coordinate system
            
    Dim swCoordSys As SldWorks.CoordinateSystemFeatureData
    Dim swMathTransform As SldWorks.MathTransform
    
    Set swCoordSys = selectedCoordsFeat.GetDefinition
    Set swMathTransform = swCoordSys.Transform
    
    Dim vMatrix As Variant
    vMatrix = swMathTransform.ArrayData

    Set swFeatMgr = swModel.FeatureManager
    
    ' Create the feature folder if it doesn't exist
    Dim swFeatFolder As SldWorks.Feature
    Set swFeatFolder = swModel.FeatureByName("Sensor coordinates")
    
    If swFeatFolder Is Nothing Then
        Set swFeatFolder = swFeatMgr.InsertFeatureTreeFolder2(swFeatureTreeFolder_EmptyBefore)
        swFeatFolder.Name = "Sensor coordinates"
    End If
    
    Dim angX, angY, angZ As Double
    Dim mirrorCoordsFeat  As SldWorks.Feature
    
    ' Computations are made according to Tait-Bryan ZYX angles
    ' The X-coordinates of the XYZ vectors are inverted to mirror about YZ plane
    ' The X vector (1st row of vMatrix) is reversed by inverting all components
    
    Const pi As Double = 3.14159265
    
    If vMatrix(0) = 0 And vMatrix(3) = 0 Then
        angZ = 0
        angY = -pi / 2
        angX = -ArcTan2(CDbl(-vMatrix(1)), CDbl(vMatrix(4)))
    Else
        angZ = -ArcTan2(CDbl(-vMatrix(3)), CDbl(vMatrix(0)))
        angY = -ArcTan2(CDbl(vMatrix(6)), Sqrt(CDbl(vMatrix(0)) ^ 2 + CDbl(-vMatrix(3)) ^ 2))
        angX = -ArcTan2(CDbl(vMatrix(7)), CDbl(vMatrix(8)))
    End If
    
    Set mirrorCoordsFeat = swFeatMgr.CreateCoordinateSystemUsingNumericalValues(True, -vMatrix(9), vMatrix(10), vMatrix(11), True, angX, angY, angZ)
    mirrorCoordsFeat.Name = selectedCoordsFeat.Name + "-Mirror"
    
    ' Move both coordinate systems to the created folder
    
    retVal = swModel.Extension.ReorderFeature(mirrorCoordsFeat.Name, swFeatFolder.Name, swMoveToFolder)
    retVal = swModel.Extension.ReorderFeature(selectedCoordsFeat.Name, swFeatFolder.Name, swMoveToFolder)
End Sub