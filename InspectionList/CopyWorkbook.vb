Sub OpenWorkbook(Filename As String)
    MsgBox (Filename)
    Application.ScreenUpdating = False
    
    Dim src As Workbook
    
    ' OPEN THE SOURCE EXCEL WORKBOOK IN "READ ONLY MODE".
    Set src = Workbooks.Open(Filename, True, True)
    
    ' GET THE TOTAL ROWS FROM THE SOURCE WORKBOOK BY GETTING ALL ROWS IN COLUMN B.
    Dim TotalRows As Integer
    TotalRows = src.Worksheets("sheet1").Range("B1:B" & Cells(Rows.Count, "B").End(xlUp).Row).Rows.Count
    
    MsgBox (TotalRows)
   
    ' COPY DATA FROM SOURCE (CLOSE WORKGROUP) TO THE DESTINATION WORKBOOK.
    Dim i As Integer            ' COUNTER.
    For i = 1 To TotalRows
        Worksheets("Data").Range("B" & i).Formula = src.Worksheets("Sheet1").Range("B" & i).Formula
    Next i
    
    ' CLOSE THE SOURCE FILE.
    src.Close False             ' FALSE - DON'T SAVE THE SOURCE FILE.
    Set src = Nothing
    
ErrHandler:
    Application.EnableEvents = True
    Application.ScreenUpdating = True
End Sub