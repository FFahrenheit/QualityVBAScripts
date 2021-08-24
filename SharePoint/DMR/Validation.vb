Sub SaveAlert()
    On Error GoTo DebugFunction
    Dim I As Integer
    Length = Worksheets("Validation").Range("A2", Worksheets("Validation").Range("A2").End(xlDown)).Rows.Count
    MsgBox Length
    
    For I = 2 To Length + 1
    
        Reference = Worksheets("Validation").Range("A" & I)
        Value = Range(Reference)
        
        Content = "Por favor, " & Worksheets("Validation").Range("B" & I)
        Title = Worksheets("Validation").Range("C" & I)
        
        If Value = "" Then
            CallNotification Title, Content, Reference
            Exit Sub
        End If
        
    Next I
    
    MsgBox "Hola!!!"
    Exit Sub
DebugFunction:
    MsgBox Reference
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

