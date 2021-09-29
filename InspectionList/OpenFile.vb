Function GetFile() As String 
    Dim SeletedFile As String
    SelectedFile = ""

    With Application.FileDialog(msoFileDialogFilePicker)        
        .ButtonName = "Confirm"
        .AllowMultiSelect = False
        .Title = "Seleccionar el archivo generado"
        .Filters.Clear
        .Filters.Add "Excel Worksheets", "*.xls; *.xlsx; *.xlsm"
        .InitialFileName = "D:\"
        If .Show = -1 Then
        'ok
            SelectedFile = .SelectedItems(1)
            GetFile = SelectedFile
            ' OpenWorkbook (SelectedFile)
        Else
        'cancel
            MsgBox "No se pudieron cargar los datos de la hoja de inspeccion", vbOKOnly + vbCritical, "Error de carga"
        End If
        
    End With
End Function

Sub GettingFile()
    Filename = GetFile()
    If Filename <> "" Then
        OpenWorkbook (SelectedFile)
    End If 
End Sub

Sub LoadTemplate()
    Filename = GetFile()
    If Filename <> "" Then
        CopyTemplate (Filename)
    End If 
End Sub