Sub CopyTemplate(Filename As String)
    Dim A As String
    Dim C As String
    
    A = ActiveWorkbook.Name  'Workbook destino
    C = "Diccionario"
    
    Exists = SheetExists(C, A)
    If Exists = True Then
        'Existe la worksheet
    Else
        'No existe, la creamos
        With Workbooks(A)
            Set ws = .Sheets.Add(After:=.Sheets(.Sheets.Count))
            ws.Name = C
        End With
    End If
    
    
    
    Dim Destino As Worksheet
    Dim Dict As Worksheet
    Dim src As Workbook
    
    Application.ScreenUpdating = False
    'MsgBox "Filename: " & Filename
    
    Set src = Workbooks.Open(Filename, True, True)
    
    Y = src.Name
    Z = "Sheet1"
    
    Set Destino = Workbooks(Y).Worksheets(Z)
    Set Dict = Workbooks(A).Worksheets(C)
        
    'Header details
    Dict.Range("A1") = "Pieza"
    Dict.Range("A2") = "Fecha"
    Dict.Range("A3") = "Hora"

    Dict.Range("B1") = "N/A"
    Dict.Range("B2") = 14
    Dict.Range("B3") = 21
    
    i = 10
    j = 5
    
    While Destino.Range("B" & i) <> ""
        Cota = Destino.Range("B" & i)
        Dict.Range("A" & j) = Cota
        
        i = i + 4
        j = j + 1
    Wend
    
    'Close all
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    src.Close (False)
    Set src = Nothing
    MsgBox "Temp ok"
    Exit Sub
ErrHandler:
    MsgBox "Error: " & Err.Description
    MsgBox "No se pudo cargar la plantilla de la hoja de inspeccion", vbOKOnly + vbCritical, "Error de carga"
    Exit Sub
End Sub