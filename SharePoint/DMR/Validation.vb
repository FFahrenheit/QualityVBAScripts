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
        If Range("fiveWhy") = "" Then
            CallNotification "FALTA LLENAR 5 POR QUÉ", "Por favor, llene al menos uno de los 5 por qué y siga el orden del diagrama", "fiveWhy"
            Exit Sub
        End If
        
        I = 21
        Filled = False
        
        While I <= 25
            Action = Worksheets("DMR Hoja 2").Range("A" & I)
            Responsable = Worksheets("DMR Hoja 2").Range("G" & I)
            If Action = "" Xor Responsable = "" Then
                CallNotification "FALTA LLENAR PLAN DE ACCIÓN", "Por favor, llene todas las acciones que están incompletas", "action"
                Exit Sub
            ElseIf Action <> "" And Responsable <> "" Then
                Filled = True
            End If
            I = I + 1
        Wend
        
        If Filled = False Then
            CallNotification "FALTA LLENAR PLAN DE ACCIÓN", "Por favor, describa al menos una acción para evitar el problema", "action"
            Exit Sub
        End If
        
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

