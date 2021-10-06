' Recover excel app when everything goes wrong
Sub Recover()
    Application.DisplayFormulaBar = True
    Application.DisplayStatusBar = True
    For Each Worksheet In ActiveWorkbook.Worksheets
        Worksheet.Unprotect ("Calidad2020")
        Worksheet.Visible = True
    Next Worksheet
    ActiveWindow.DisplayHeadings = True
    ActiveWindow.DisplayWorkbookTabs = True
    MsgBox "Application recovery successful"
End Sub

' Check if sheet exists
Function SheetExists(sheetToFind As String, Optional WB As String) As Boolean
    If WB = "" Then
        Set InWorkbook = ThisWorkbook
    Else
        Set InWorkbook = Workbooks(WB)
    End If

    Dim Sheet As Object
    For Each Sheet In InWorkbook.Sheets
        If sheetToFind = Sheet.Name Then
            SheetExists = True
            Exit Function
        End If
    Next Sheet
    SheetExists = False
End Function

'' This functions opens a file explorer dialog and
'' return the string of the path to the file (excel files)
Function GetFile() As String
    Dim SeletedFile As String
    SelectedFile = ""

    With Application.FileDialog(msoFileDialogFilePicker)
        .ButtonName = "Confirmar"
        .AllowMultiSelect = False
        .Title = "Seleccionar el archivo generado"
        .Filters.Clear
        .Filters.Add "Excel Worksheets", "*.xls; *.xlsx; *.xlsm"
        .InitialFileName = "D:\"
        If .Show = -1 Then
        'ok
            SelectedFile = .SelectedItems(1)
            GetFile = AreYouSure((SelectedFile))
            ' OpenWorkbook (SelectedFile)
        Else
        'cancel
            MsgBox "No se pudieron cargar los datos de la hoja de inspeccion", vbOKOnly + vbCritical, "Error de carga"
        End If
        
    End With
End Function

Function AreYouSure(Filename As String) As String
    Msg = "Ha seleccionado el archivo " & Filename & vbCrLf _
        & "¿Está seguro de continuar? (Seleccione No para cambiar de archivo)"
    Response = MsgBox(Msg, vbYesNoCancel + vbQuestion, "CONFIRMACION")
    
    If Response = vbYes Then
        AreYouSure = Filename
    ElseIf Response = vbNo Then
        AreYouSure = GetFile()
    Else
        AreYouSure = ""
    End If
    
End Function
