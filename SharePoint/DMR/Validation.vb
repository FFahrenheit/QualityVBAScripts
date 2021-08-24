Sub SaveAlert()
    On Error GoTo Fault
    Result = Validate()
    'MsgBox "Result = " & Result
    
    If Result = False Then
        Exit Sub
    End If
    
    Message = "El DMR será guardado en el SharePoint de QMS en la libreria de <libreria>.  " _
        & " Aqui podra consultarla para posterior revisión y/o edición." _
        & "¿Desea continuar?"
    
    Style = vbYesNo + vbExclamation + vbDefaultButton2
    
    Response = MsgBox(Message, Style, "SALVAR ARCHIVO")
    
    If Response = vbYes Then
    
        Filename = "DMR-" & Format(Range("number"), "000#")
        
        ActiveWorkbook.SaveAs Filename:= _
        "https://interplexgroup.sharepoint.com/americas/imx/imx_qms/TEMPLATES/" & Filename & ".xlsm" _
        , FileFormat:=xlOpenXMLWorkbookMacroEnabled, Password:="", WriteResPassword:="", _
        ReadOnlyRecommended:=False, CreateBackup:=False
        

        Msg = "El DMR ha sido guardado con éxito"
        Style = vbOKOnly + vbInformation
        
        MsgBox Msg, Style, "DMR GUARDADO"
        
    End If
    
    Exit Sub

Fault:
    Msg = "El DMR no puede ser grabado, por favor" _
        & " asegurese que no haber utilizado caracteres especiales en el nombre como " _
        & "(!?¡*[]&$()%@/), esto evita que el DMR pueda ser guardada, " _
        & "los guiones bajos ( _ ) SI pueden ser utilizados" & Chr(10) & Chr(10) & "En caso contrario por favor " _
        & "informe al administrador del sharepoint sobre el problema"
    Style = vbOKOnly + vbInformation
    
    MsgBox Msg, Style, "ERROR AL GUARDAR"
End Sub


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
        
        If Value = "" Then
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
    
    If Not (Range("production") Or Range("warehouse") Or Range("planning")) Then
        CallNotification "FALTA AVISAR MOVIMIENTO DE MATERIALES", "Por favor, indique al menos a un área que se avisó del movimiento de materiales", "production"
        Exit Function
    End If
    
    Risk = Range("riskLevel")
    
    Reasons = Range("noRiskReasons")
    If Risk = "No" And Reasons = "" Then
        CallNotification "FALTA LLENAR RAZONES", "Por favor, llene la razón y acciones de porqué no hay riesgo", "noRiskReasons"
        Exit Function
    End If
    
    If Risk = "Medio" Or Risk = "Alto" Then
        If Range("fiveWhy") = "" Then
            CallNotification "FALTA LLENAR 5 POR QUÉ", "Por favor, llene al menos uno de los 5 por qué y siga el orden del diagrama de la hoja 2", "fiveWhy"
            Exit Function
        End If
        
        I = 21
        Filled = False
        
        While I <= 25
            Action = Worksheets("DMR Hoja 2").Range("A" & I)
            Responsable = Worksheets("DMR Hoja 2").Range("G" & I)
            If Action = "" Xor Responsable = "" Then
                CallNotification "FALTA LLENAR PLAN DE ACCIÓN", "Por favor, llene todas las acciones que están incompletas en la hoja 2", "action"
                Exit Function
            ElseIf Action <> "" And Responsable <> "" Then
                Filled = True
            End If
            I = I + 1
        Wend
        
        If Filled = False Then
            CallNotification "FALTA LLENAR PLAN DE ACCIÓN", "Por favor, describa al menos una acción para evitar el problema en la hoja 2", "action"
            Exit Function
        End If
        
    End If
    
    Validate = True
    Exit Function
DebugFunction:
    MsgBox "Error: " & Err.Description
End Function

Sub CallNotification(Title, Content, Cell)
    On Error GoTo EndFunction
    Style = vbOKOnly + vbCritical
    Response = MsgBox(Content, Style, Title)
    
    If Response = vbOK Then
        Range(Cell).Select
    End If
    
EndFunction:
End Sub

