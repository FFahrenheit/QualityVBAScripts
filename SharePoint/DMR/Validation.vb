Sub SaveAlert()
    Range("Number").Select
    
    Dim I As Integer
    Length = Range("A2", Range("A2").End(xlDown)).Rows.Count
    
    For I = 2 To Length + 1
    
        Reference = Range("Validation!A" & I)
        Value = Range(Reference)
        Content = "Por favor, " & Range("Validation!B" & I)
        Title = Range("Validation!C" & I)
        
        If Value = "" Then
            CallNotification Title, Content, Reference
            Exit For
        End If
        
    Next I
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

