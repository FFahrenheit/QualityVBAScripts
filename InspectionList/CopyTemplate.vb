Sub LoadTemplate()
    Filename = GetFile()
    If Filename <> "" Then
        CopyTemplateA1 (Filename)
    End If
End Sub
 

Sub CopyTemplateA1(Filename As String)
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
    
    Set Destino = Workbooks(Y).Sheets(1)
    Set Dict = Workbooks(A).Worksheets(C)
        
    Dict.Columns(1).ClearContents
    'Header details
    Dict.Range("A1") = "Pieza"
    Dict.Range("A2") = "Fecha"
    Dict.Range("A3") = "Hora"

    Dict.Range("B1") = "N/A"
    Dict.Range("B2") = 14
    Dict.Range("B3") = 21
    
    I = 10
    J = 5
    
    While Destino.Range("B" & I) <> ""
        Cota = Destino.Range("B" & I)
        Dict.Range("A" & J) = Cota
        
        I = I + 4
        J = J + 1
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

Sub CopyTemplateB(Filename As String)
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
    
    Set Destino = Workbooks(Y).Sheets(1)
    Set Dict = Workbooks(A).Worksheets(C)
    
    Dict.Columns(1).ClearContents
    'Header details
    Dict.Range("A1") = "Pieza"
    Dict.Range("A2") = "Fecha"
    Dict.Range("A3") = "Hora"

    Dict.Range("B1") = "N/A"
    Dict.Range("B2") = 14
    Dict.Range("B3") = 21
    
    I = 2 'Fila
    J = 5
    K = 1 'Columna
    
    Do While IsNumeric(Destino.Cells(2, K)) = True
        Cota = Destino.Cells(1, K) & "[" & (I - 1) & "]"
        Dict.Range("A" & J) = Cota
        
        I = I + 1
        J = J + 1
        If Destino.Cells(I, K) = "" Then
            I = 2
            K = K + 1
        End If
    Loop
    
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

Sub CopyTemplateA3(Filename As String)
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
    
    Set Destino = Workbooks(Y).Sheets(1)
    Set Dict = Workbooks(A).Worksheets(C)
        
    Dict.Columns(1).ClearContents
    'Header details
    Dict.Range("A1") = "Pieza"
    Dict.Range("A2") = "Fecha"
    Dict.Range("A3") = "Hora"

    Dict.Range("B1") = "N/A"
    Dict.Range("B2") = 14
    Dict.Range("B3") = 21
    
    I = 10
    J = 5
    
    While Destino.Range("B" & I) <> ""
        I = I + 2
        
        Do While Destino.Range("B" & I) <> ""
            Cota = Destino.Range("B" & I) & " " & Destino.Range("E" & I)
            Dict.Range("A" & J) = Cota
            I = I + 1
            J = J + 1
        Loop
        
        I = I + 1
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



