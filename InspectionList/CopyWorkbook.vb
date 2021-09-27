Sub OpenWorkbook(Filename As String)
    'On Error GoTo ErrHandler
    A = ActiveWorkbook.Name  'Workbook destino
    B = "Data"               'Hoja destino
    
    Dim Destino As Worksheet
    Dim Origen As Worksheet
    Dim src As Workbook
    
    Application.ScreenUpdating = False
    MsgBox "Filename: " & Filename
    
    Set src = Workbooks.Open(Filename, True, True)
    Y = src.Name            'Workbook origen
    Z = "Sheet1"            'Hoja origen
    
    
    Set Origen = Workbooks(A).Worksheets(B)
    Set Destino = Workbooks(Y).Worksheets(Z)
    
    Origen.Range("B2").Value = Destino.Range("B2")

    
    
    'Close all
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    src.Close (False)
    Set src = Nothing
    Exit Sub
ErrHandler:
    MsgBox "Error: " & Err.Description
    Exit Sub
End Sub