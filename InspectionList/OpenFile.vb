Sub GetingFile()
    Dim SelectedFile As String
    
    With Application.FileDialog(msoFileDialogFilePicker)
        
        .ButtonName = "Confirm"
        .AllowMultiSelect = False
        .Title = "Seleccionar el archivo generado"
        .Filters.Clear
        .Filters.Add "Extensible Markup Language ", "*.xml"
        .InitialFileName = "D:\"
        If .Show = -1 Then
        'ok'
            SelectedFile = .SelectedItems(1)
            MsgBox (SelectedFile)
        Else
        'cancel'
        End If
        
    End With
End Sub