Sub GettingFile()
    Dim SelectedFile As String
    
    With Application.FileDialog(msoFileDialogFilePicker)
        
        .ButtonName = "Confirm"
        .AllowMultiSelect = False
        .Title = "Seleccionar el archivo generado"
        .Filters.Clear
        .Filters.Add "Excel Worksheets", "*.xls; *.xlsx"
        .InitialFileName = "D:\"
        If .Show = -1 Then
        'ok'
            SelectedFile = .SelectedItems(1)
            OpenWorkbook (SelectedFile)
        Else
        'cancel'
        End If
        
    End With
End Sub