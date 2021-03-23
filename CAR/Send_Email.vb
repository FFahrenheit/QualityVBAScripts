Sub Send_Email()
    On Error GoTo ErrHandler
    'Get data'
    Worksheets("Summary").Activate

    LastUpdate = Range("F2")
    CurrentDate = Date
    
    If IsEmpty(LastUpdate) Or DateDiff("d", LastUpdate, CurrentDate) >= 1 Then
        NumRows = Range("B3", Range("B3").End(xlDown)).Rows.Count
        
        Dim html
        html = "<!DOCTYPE html><html><body>"
        html = html & "<div style=""font-family:'Segoe UI', Calibri, Arial, Helvetica; font-size: 14px; max-width: 768px;"">"
        html = html & "Buen día, equipo. <br /><br /> Envio el estado actual de la bitácora de CAR <br />"
        html = html & "Aquí las próximas fechas de documentacion y documentación vencida<br /><br />"
        html = html & "<table style='border-spacing: 0px; border-style: solid; border-color: #ccc; border-width: 0 0 1px 1px;'>"
    
        Range("B3").Select
    
        For I = 1 To NumRows
            html = html & "<tr>"
            Header = Range("B" & I + 2)
            Value = Range("C" & I + 2)
    
            html = html & "<td style='padding: 10px; border-style: solid; border-color: #ccc; border-width: 1px 1px 0 0;'>" & Header & "</td>"
            html = html & "<td style='padding: 10px; border-style: solid; border-color: #ccc; border-width: 1px 1px 0 0;'>" & Value & "</td>"
    
            html = html & "</tr>"
        Next
    
        html = html & "</table></div></body></html>"
    
        Dim objOutlook As Object
        Set objOutlook = CreateObject("Outlook.Application")
        
    
        Dim objEmail As Object
        Set objEmail = objOutlook.CreateItem(olMailItem)
    
        With objEmail
            .To = "i.lopez@mx.interplex.com;Martha.Rodriguez@mx.interplex.com"
            .Subject = "Estado Bitácora CAR"
            .HTMLBody = html
            .Send
        End With
        
        Set objEmail = Nothing:    Set objOutlook = Nothing
        Range("F2").Value = CurrentDate
        
        Exit Sub
    End If

ErrHandler:
    Range("F2").Value = LastUpdate

End Sub
