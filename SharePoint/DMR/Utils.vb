
Sub CallNotification(Title, Content, Cell)
    On Error GoTo EndFunction
    Style = vbOKOnly + vbCritical
    Response = MsgBox(Content, Style, Title)
    
    If Response = vbOK Then
        Range(Cell).Select
    End If
    
EndFunction:
End Sub

'Recorded macro!
Sub AddAction()
    Rows("22:22").Select
    Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("A21:T21").Select
    Selection.Copy
    Range("A22").Select
    ActiveSheet.Paste
    Range("A22:T22").Select
    Application.CutCopyMode = False
    Selection.ClearContents
    Range("A22:F22").Select
    ActiveCell.FormulaR1C1 = ""
    Range("A22:F22").Select
    
End Sub