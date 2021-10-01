Sub Recover()
    Application.DisplayFormulaBar = True
    Application.DisplayStatusBar = True
    For Each Worksheet In ActiveWorkbook.Worksheets
        Worksheet.Unprotect ("Calidad2020")
        Worksheet.Visible = True
    Next Worksheet
    ActiveWindow.DisplayHeadings = True
    ActiveWindow.DisplayWorkbookTabs = True
End Sub
