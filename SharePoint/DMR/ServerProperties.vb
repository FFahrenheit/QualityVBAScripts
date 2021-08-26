Sub SetServerProperties()
    Dim WB As Workbook
    Set WB = ThisWorkbook
    
    For Each Prop In WB.ContentTypeProperties
        On Error Resume Next
        'MsgBox Prop.Name
        Select Case Prop.Name
            Case "#DMR"
                Prop.Value = Range("dmrNumber").Value
            Case "Cliente"
                Prop.Value = Range("clientName").Value
            Case "IMX_PN"
                Prop.Value = Range("partNumber").Value
            Case "Fecha_Creado"
                Prop.Value = Range("date").Value
            Case "Defecto"
                Prop.Value = Range("defective").Value
            Case "Discrepancia"
                Prop.Value = Range("discrepancy").Value
            Case "Originador"
                Prop.Value = Range("originator").Value
            Case "Disposición"
                Prop.Value = Range("disposition").Value
            Case "Lote"
                Prop.Value = Range("lotSize").Value
            Case "Area_Rechazo"
                Prop.Value = Range("area").Value
            Case "Responsable_Validar_Purga"
                Prop.Value = Range("validator").Value
            Case "Numero_Orden"
                Prop.Value = Range("orderNumber").Value
            Case "Riesgo_Repeticion"
                Prop.Value = Range("riskLevel").Value
            Case "Lider_Linea"
                Prop.Value = Range("lineaLeader").Value
            Case "Muestra"
                Prop.Value = Range("lot").Value
            Case "Numero_ID"
                'MsgBox "Es el ID!!"
                If Range("number").Value = "" Then
                    Prop.Value = "9999"
                Else
                    Prop.Value = Range("number").Value
                End If
            Case Else
                'N/A
        End Select
    Next Prop
End Sub