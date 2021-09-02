Private Sub Worksheet_Change(ByVal Target As Range)
    ActiveSheet.Activate
    'MsgBox Target.Value
    If IsCell(Target, "riskLevel") Then
        If Target.Value <> "No" Then
            Rows("36:36").EntireRow.Hidden = True
            Range("noRiskReasons").Value = ""
        Else
            Rows("36:36").EntireRow.Hidden = False
        End If
        
    ElseIf IsCell(Target, "disposition") Then
        If Target.Value <> "Usar así" Then
            Rows("31:31").EntireRow.Hidden = True
        Else
            Rows("31:31").EntireRow.Hidden = False
        End If
    End If
End Sub

Function IsCell(Target As Range, Name As String) As Boolean
    IsCell = Not Application.Intersect(Range(Name), Range(Target.Address)) Is Nothing
End Function
