Function SheetExists(sheetToFind As String, Optional WB As String) As Boolean

    If WB = "" Then
        Set InWorkbook = ThisWorkbook
    Else
        Set InWorkbook = Workbooks(WB)
    End If

    Dim Sheet As Object
    For Each Sheet In InWorkbook.Sheets
        If sheetToFind = Sheet.Name Then
            SheetExists = True
            Exit Function
        End If
    Next Sheet
    SheetExists = False
End Function