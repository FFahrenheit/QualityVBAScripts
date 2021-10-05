Sub FindBlank()
    i = 1
    'MsgBox ActiveSheet.Cells(1, 1)
    'Exit Sub
    While ActiveSheet.Cells(1, i) <> ""
        'MsgBox ActiveSheet.Cells(1, i)
        i = i + 1
    Wend
    
    MsgBox "Index = " & i
    
End Sub

Sub LastCell()
    Dim WS As Worksheet
    Dim i As Integer
    Set WS = ActiveWorkbook.Sheets("Hoja de inspeccion")
    i = 54
    
    'MsgBox WS.Cells(14, i).Value
    
    Do While True
        If ((WS.Cells(14, i).Value = "-" Or WS.Cells(14, i) = "") And WS.Cells(21, i).Value = "" _
            And WS.Cells(22, i).Value = "") Or i > 60 Then
            Exit Do
        End If
        i = i + 1
    Loop
    
    MsgBox i
End Sub