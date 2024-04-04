Sub main()

    Dim swApp As SldWorks.SldWorks
    Set swApp = Application.SldWorks
    
    Dim swModel As SldWorks.ModelDoc2
    Set swModel = swApp.ActiveDoc
    
    Dim swFeat As SldWorks.feature
    
    If Not swModel Is Nothing Then
        
        Set swSelMgr = swModel.SelectionManager
        
        ' Get the path of the open file
        Dim fpath, csvFilePath As String
        
        fpath = swModel.GetPathName
        
        ' Extract the directory portion without the filename
        fpath = Left(fpath, InStrRev(fpath, "\"))
        
        csvFilePath = fpath & "sensorcoords.csv"
        
        ' Open file for writing
        csvFileNum = FreeFile()
        Open csvFilePath For Output As csvFileNum
        
        ' CSV Title row
        Print #csvFileNum, "Slot,X,Y,Z,XX,XY,XZ,YX,YY,YZ,ZX,ZY,ZZ"
        
        ' Process each selected feature
        If Not swSelMgr.GetSelectedObjectCount2(-1) = 0 Then
            For j = 1 To swSelMgr.GetSelectedObjectCount2(-1)
                Set swFeat = swSelMgr.GetSelectedObject6(j, -1)
                
                Dim swMat As SldWorks.MathTransform
                Dim swRotMat As Variant
                
                ' Check if feature is null
                If swFeat Is Nothing Then
                    MsgBox "Invalid feature"
                    Exit Sub
                End If
                
                ' Get the rotation matrix of the feature
                Set swMat = swFeat.GetDefinition.Transform
                swRotMat = swMat.ArrayData
                
                ' Write location XYZ and local coordinate axes to CSV
                Print #csvFileNum, swFeat.Name & ",";
                Print #csvFileNum, Replace(Format(swRotMat(9), "0.0000"), ",", ".") & "," _
                                 & Replace(Format(swRotMat(10), "0.0000"), ",", ".") & "," _
                                 & Replace(Format(swRotMat(11), "0.0000"), ",", ".") & ",";
                
                For k = 0 To 8
                    If k < 8 Then
                        Print #csvFileNum, Replace(Format(swRotMat(k), "0.0000"), ",", ".") & ",";
                    ElseIf j < swSelMgr.GetSelectedObjectCount2(-1) Then
                        Print #csvFileNum, Replace(Format(swRotMat(k), "0.0000"), ",", ".")
                    Else
                        ' Prevent extra LF at the end of the file
                        Print #csvFileNum, Replace(Format(swRotMat(k), "0.0000"), ",", ".");
                    End If
                Next k
            Next j
        Else
            MsgBox "Please select coordinate system features"
        End If
        
        MsgBox "Coordinate systems saved to " & csvFilePath, vbInformation
        Close csvFileNum
        
    Else
        MsgBox "Please open model"
    End If
    
End Sub
