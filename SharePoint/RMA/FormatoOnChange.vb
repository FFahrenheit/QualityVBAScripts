Private Sub Worksheet_Change(ByVal Target As Range)
    ActiveSheet.Activate
    'MsgBox Target.Value
    If Not Application.Intersect(Range("disposition"), Range(Target.Address)) Is Nothing Then
        
        If Target.Value <> "Scrap" Then
            Rows("19:20").EntireRow.Hidden = True
            Range("scrap").Value = ""
        Else
            Rows("19:20").EntireRow.Hidden = False
        End If
        
    End If
End Sub
