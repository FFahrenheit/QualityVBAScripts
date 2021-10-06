'' This module reads the file and uploads all the measures into the
'' inpsection list

Sub LoadData()
    Dim Filename As String
    
    Message = "Esta opcion le permitirá cargar los datos de las hojas de inspeccion, " _
    & "usela si desea cargar nuevas mediciones a la hoja. ¿Desea continuar?"
    Response = MsgBox(Message, vbYesNo + vbQuestion + vbDefaultButton1, "CAMBIAR PLANTLLA")
    
    If Response = vbYes Then
        Filename = Utils.GetFile()
        If Filename <> "" Then
            ReadData (Filename)
        End If
    End If
End Sub

Sub ReadData(Filename As String)
    Dim A As String
    Dim B As String
    On Error GoTo Handler
    A = ActiveWorkbook.Name  'Workbook destino
    B = "Data"               'Hoja destino
    
    Exists = Utils.SheetExists(B, A)
    If Exists = False Then
        'No existe, la creamos
        With Workbooks(A)
            Set WS = .Sheets.Add(After:=.Sheets(.Sheets.Count))
            WS.Name = B
        End With
    End If
    
    Dim Destino As Worksheet
    Dim Origen As Worksheet
    'Dim Dict As Worksheet
    Dim src As Workbook
    
    Application.ScreenUpdating = False
    'MsgBox "Filename: " & Filename
    
    Set src = Workbooks.Open(Filename, True, True)
    
    Y = src.Name
    Z = "Sheet1"
    
    Set Origen = Workbooks(A).Worksheets(B)
    Set Destino = Workbooks(Y).Worksheets(Z)
    'Set Dict = Workbooks(A).Worksheets(C)
    
    Origen.Visible = True
    Origen.Columns(1).ClearContents
    Origen.Columns(2).ClearContents
    
    'Header details
    Origen.Range("A1") = "Pieza"
    Origen.Range("A2") = "Fecha"
    Origen.Range("A3") = "Hora"
    
    Origen.Range("B1") = Destino.Range("C3")
    Origen.Range("B2") = Destino.Range("C6")
    Origen.Range("B2").NumberFormat = "dd/mm/yyyy"
    Origen.Range("B3") = Destino.Range("C7")
    Origen.Range("B3").NumberFormat = "hh:mm"
    
    I = 10
    J = 5
    
    While Destino.Range("B" & I) <> ""
        Cota = Destino.Range("B" & I)
        Valor = Destino.Range("H" & (I + 2))
        Origen.Range("A" & J) = Cota
        Origen.Range("B" & J) = Valor
        
        'Dict.Range("A" & j) = Cota
        
        I = I + 4
        J = J + 1
    Wend
    
    'Close all
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    src.Close (False)
    Set src = Nothing
    Algorithm.GetData
    ' Destino.Visible = False
    Exit Sub
Handler:
    MsgBox "Error: " & Err.Description
    MsgBox "No se pudieron cargar los datos de la hoja de inspeccion", vbOKOnly + vbCritical, "Error de carga"
End Sub
