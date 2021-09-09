Sub SetServerProperties()
    Dim WB As Workbook
    Set WB = ThisWorkbook
    
    For Each Prop In WB.ContentTypeProperties
        On Error GoTo Fault 'Resume Next
        Name = Prop.Name
        Select Case Name
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
            Case "Muestra"
                Prop.Value = Range("lot").Value
            Case "Informado_Planeacion"
                Prop.Value = Range("planning").Value
            Case "Informado_Produccion"
                Prop.Value = Range("production").Value
            Case "Informado_Almacen"
                Prop.Value = Range("warehouse").Value
            Case "Acciones_Totales"
                Prop.Value = Range("totalActions").Value
            Case "Acciones_Hechas"
                Prop.Value = Range("completedActions").Value
            Case "Progreso_Acciones"
                Prop.Value = Range("progress").Value
            Case "Turno_Creado"
                Prop.Value = Range("shift").Value
            Case "Ultima_Accion"
                Prop.Value = Range("lastAction").Value
            Case "Acciones_Completas"
                Prop.Value = Range("actionsCompleted").Value
            Case "Acciones_Disponibles"
                Prop.Value = Range("actionsAvailable").Value
            Case "5Porque_Disponible"
                Prop.Value = Range("whyAvailable").Value
            Case "Problema_Identificado"
                Prop.Value = Range("problem").Value
            Case "Numero_ID"
                If Range("number").Value = "" Then
                    Prop.Value = "9999"
                Else
                    Prop.Value = Range("number").Value
                End If
            Case Else
                'N/A
        End Select
    Next Prop
    Exit Sub
Fault:
    MsgBox Name
End Sub