'On workbook
Private Sub Workbook_BeforeSave(ByVal SaveAsUI As Boolean, Cancel As Boolean)
    Call SetServerProperties
End Sub

'On Module
Sub SetServerProperties()
    Dim WB As Workbook
    Set WB = ThisWorkbook
    
    For Each Prop In WB.ContentTypeProperties
        On Error Resume Next
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
                Prop.Value = Range("lineLeader").Value
            Case "Muestra"
                Prop.Value = Range("lot").Value
            Case Else
                'N/A
        End Select
    Next Prop
End Sub