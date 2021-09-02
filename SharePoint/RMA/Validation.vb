Function Validate() As Boolean
    Validate = False
    On Error GoTo DebugFunction
    Dim I As Integer
    Length = Worksheets("Validacion").Range("A2", Worksheets("Validacion").Range("A2").End(xlDown)).Rows.Count
    
    For I = 2 To Length + 1
        Reference = Worksheets("Validacion").Range("A" & I)
        Value = Range(Reference)
        
        Content = "Por favor, " & Worksheets("Validacion").Range("B" & I)
        Title = Worksheets("Validacion").Range("C" & I)
        
        If Value = "" Then
            CallNotification Title, Content, Reference
            Exit Function
            'Exit For
        End If
    Next I

    If Range("disposition") = "Scrap" And Range("scrap") = "" Then
        CallNotification "FALTA LLENADO", "Por favor, seleccione donde se hará el Scrap", "scrap"
        Exit Function
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
