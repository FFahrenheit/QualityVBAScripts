Sub CallNotification(Title, Content, Cell)
    On Error GoTo EndFunction
    Style = vbOKOnly + vbCritical
    Response = MsgBox(Content, Style, Title)
    
    If Response = vbOK Then
        Range(Cell).Select
    End If
    
EndFunction:
End Sub

Sub AddAction()
    Dim WS As Worksheet
    WB = ActiveWorkbook.Name
    WSName = ActiveSheet.Name
    Set WS = Workbooks(WB).Sheets(WSName)
    WS.Rows("22:22").Select
    Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
    WS.Range("A21:T21").Select
    Selection.Copy
    WS.Range("A22").Select
    ActiveSheet.Paste
    WS.Range("A22:T22").Select
    Application.CutCopyMode = False
    Selection.ClearContents
    WS.Range("A22:F22").Select
    ActiveCell.FormulaR1C1 = ""
    WS.Range("A22:F22").Select
End Sub
