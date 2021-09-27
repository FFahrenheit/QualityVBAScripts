Sub OpenWorkbook(Filename As String)
    'On Error GoTo ErrHandler
    A = ActiveWorkbook.Name  'Workbook destino
    B = "Data"               'Hoja destino
    C = "Diccionario"
    
    Dim Destino As Worksheet
    Dim Origen As Worksheet
    Dim Dict As Worksheet
    Dim src As Workbook
    
    Application.ScreenUpdating = False
    'MsgBox "Filename: " & Filename
    
    Set src = Workbooks.Open(Filename, True, True)
    
    Y = src.Name
    Z = "Sheet1"
    
    Set Origen = Workbooks(A).Worksheets(B)
    Set Destino = Workbooks(Y).Worksheets(Z)
    Set Dict = Workbooks(A).Worksheets(C)
    
    
    'Header details
    Origen.Range("B1") = Destino.Range("C3")
    Origen.Range("B2") = Destino.Range("C6")
    Origen.Range("B3") = Destino.Range("C7")
    
    i = 10
    j = 5
    
    While Destino.Range("B" & i) <> ""
        Cota = Destino.Range("B" & i)
        Valor = Destino.Range("H" & (i + 2))
        Origen.Range("A" & j) = Cota
        Origen.Range("B" & j) = Valor
        
        Dict.Range("A" & j) = Cota
        
        i = i + 4
        j = j + 1
    Wend
    
    'Close all
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    src.Close (False)
    Set src = Nothing
    
    MsgBox "Datos cargados con éxito"
    Exit Sub
ErrHandler:
    MsgBox "Error: " & Err.Description
    MsgBox "No se pudieron cargar los datos de la hoja de inspeccion"
    Exit Sub
End Sub
