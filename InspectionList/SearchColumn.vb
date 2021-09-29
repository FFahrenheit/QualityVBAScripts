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