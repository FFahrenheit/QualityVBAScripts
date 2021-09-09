Function Validate() As Boolean
    Validate = False
    On Error GoTo DebugFunction
    Dim I As Integer
    Length = Worksheets("Validation").Range("A2", Worksheets("Validation").Range("A2").End(xlDown)).Rows.Count
    
    For I = 2 To Length + 1
        Reference = Worksheets("Validation").Range("A" & I)
        Value = Range(Reference)
        
        Content = "Por favor, " & Worksheets("Validation").Range("B" & I)
        Title = Worksheets("Validation").Range("C" & I)
        
        If Value = "" Or (Value = "9999" And Reference = "number") Then
            CallNotification Title, Content, Reference
            Exit Function
            'Exit For
        End If
    Next I
    
    Contention = Range("filledActions")
    
    If Contention <> 42 Then
        CallNotification "FALTA LLENADO", "Por favor, llene toda la tabla de acciones de contención, si un valor no aplica, escriba N/A", "A17"
        Exit Function
    End If
    
'    I = 21
'    Filled = False
'
'    While Worksheets("DMR Hoja 2").Range("A" & I).Value <> Range("endActions").Value
'        Action = Worksheets("DMR Hoja 2").Range("A" & I)
'        Responsable = Worksheets("DMR Hoja 2").Range("G" & I)
'        If Action = "" Xor Responsable = "" Then
'            CallNotification "FALTA LLENAR PLAN DE ACCIÓN", "Por favor, llene todas las acciones que están incompletas en la hoja 2", "action"
'            Exit Function
'        ElseIf Action <> "" And Action <> "" Then
'            Filled = True
'        End If
'        I = I + 1
'    Wend
'
'    If I = 21 Or Filled = False Then
'        CallNotification "FALTA LLENAR PLAN DE ACCIÓN", "Por favor, llene al menos una accion a tomar en la hoja 2", "action"
'        Exit Function
'    End If
    
    
    Validate = True
    Exit Function
DebugFunction:
    MsgBox "Error: " & Err.Description
End Function