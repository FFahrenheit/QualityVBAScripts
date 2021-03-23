Sub Send_Email()
    On Error GoTo ErrHandler
    'Get data' 
    Worksheets("Summary").Activate

    NumRows = Range("B3", Range("B3").End(xlDown)).Rows.Count
    
    Dim html
    html = "<!DOCTYPE html><html><body>"
    html = html & "<div style=""font-family:'Segoe UI', Calibri, Arial, Helvetica; font-size: 14px; max-width: 768px;"">"
    html = html & "Buen dia, Calidad <br /><br /> Envio el estado actual de la bitacora de CAR <br />"
    html = html & "Here is sheet1 data:<br /><br />"
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
        .to = "i.lopez@mx.interplex.com"
        .Subject = "Estado Bitacora CAR"
        .Body = html
        .Send
    End With
    
    ' CLEAR.
    Set objEmail = Nothing:    Set objOutlook = Nothing
        
ErrHandler:
    MsgBox Err.Description
End Sub
