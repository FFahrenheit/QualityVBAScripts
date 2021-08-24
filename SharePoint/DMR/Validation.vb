Sub SaveAlert()
    On Error GoTo DebugFunction
    Dim I As Integer
    Length = Worksheets("Validation").Range("A2", Worksheets("Validation").Range("A2").End(xlDown)).Rows.Count
    
    For I = 2 To Length + 1
        Reference = Worksheets("Validation").Range("A" & I)
        Value = Range(Reference)
        
        Content = "Por favor, " & Worksheets("Validation").Range("B" & I)
        Title = Worksheets("Validation").Range("C" & I)
        
        If Value = "" Then
            CallNotification Title, Content, Reference
            'Exit Sub
            Exit For
        End If
    Next I
    
    Contention = Range("filledActions")
    
    If Contention <> 42 Then
        CallNotification "FALTA LLENADO", "Por favor, llene toda la tabla de acciones de contención, si un valor no aplica, escriba N/A", "A17"
        Exit Sub
    End If
    
    If Not (Range("production") Or Range("warehouse") Or Range("planning")) Then
        CallNotification "FALTA AVISAR MOVIMIENTO DE MATERIALES", "Por favor, indique al menos a un área que se avisó del movimiento de materiales", "production"
        Exit Sub
    End If
    
    Risk = Range("riskLevel")
    
    Reasons = Range("noRiskReasons")
    
    If Risk = "No" And Reasons = "" Then
        CallNotification "FALTA LLENAR RAZONES", "Por favor, llene la razón y acciones de porqué no hay riesgo", "noRiskReasons"
        Exit Sub
    End If
    
    If Risk = "Medio" Or Risk = "Alto" Then
        'Validación del plan de acción
    End If
    
    MsgBox "Form valido"
    
    Exit Sub
DebugFunction:
    MsgBox "Error: " & Err.Description
End Sub

Sub CallNotification(Title, Content, Cell)
    On Error GoTo EndFunction
    Style = vbOKOnly + vbCritical
    Response = MsgBox(Content, Style, Title)
    
    If Response = vbOK Then
        Range(Cell).Select
    End If
    
EndFunction:
End Sub

