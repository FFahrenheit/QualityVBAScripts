Private Sub Worksheet_Change(ByVal Target As Range)
    ActiveSheet.Activate
    If IsCell(Target, "disposition") Then
        
        If Target.Value <> "Scrap" Then
            Rows("19:20").EntireRow.Hidden = True
            Range("scrap").Value = ""
        Else
            Rows("19:20").EntireRow.Hidden = False
        End If

    ElseIf IsCell(Target, "charges") Then
        
        If Target.Value <> "Si" Then
            Rows("23:23").EntireRow.Hidden = True
            Range("poreturn").Value = "N/A"
        Else
            Rows("23:23").EntireRow.Hidden = False
            Range("poreturn").Value = "Pendiente"
        End If
    End If
End Sub

Function IsCell(Target As Range, Name As String) As Boolean
    IsCell = Not Application.Intersect(Range(Name), Range(Target.Address)) Is Nothing
End Function

