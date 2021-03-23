Sub Fill_Info()
    'Select the data'
    Worksheets("Raw Data").Activate
    Dim I As Integer
    NumRows = Range("A2", Range("A2").End(xlDown)).Rows.Count
    
    Range("A2").Select
    
    CurrentDate = Date
    
    Late2D = 0
    Upcoming4D = 0
    Late4D = 0
    Upcoming8D = 0
    Late8D = 0
    
    For I = 1 To NumRows
        Sdate = Range("M" & I + 1)
        
        CurrentStatus = Range("AR" & I + 1)
        
        'Filter empty CAR'
        If IsEmpty(Sdate) = False And CurrentStatus = "Open" Then
            CreationDate = Range("M" & I + 1)

            Date2D = Range("AM" & I + 1)
            Date4D = Range("AN" & I + 1)
            Date8D = Range("AO" & I + 1)

            DateOffset = DateDiff("d",CurrentDate,CreationDate)
            
            If IsEmpty(Date2D) Then
                If DateOffset > 1 Then
                    Late2D = Late2D + 1
                End If
            End If

            If IsEmpty(Date4D) Then
                If DateOffset >= 3 And DateOffset <= 5 Then
                    Upcoming4D = Upcoming4D +  1
                ElseIf DateOffset > 5 Then 
                    Late4D = Late4D + 1
                End If
            End If

            If IsEmpty(Date8D) Then 
                If DateOffset >= 27 And DateOffset <= 30 Then
                    Upcoming8D = Upcoming4D + 1
                ElseIf DateOffset > 30
                    Late8D + Late8D + 1 
                End If 
            End If
            
        End If
    Next
    
    MsgBox Late2D
End Sub

