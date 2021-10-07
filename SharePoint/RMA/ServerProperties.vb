Sub SetServerProperties()

    'MsgBox Format(Range("number"), "000#")
    Dim WB As Workbook
    Set WB = ThisWorkbook
    
    For Each Prop In WB.ContentTypeProperties
        On Error Resume Next
        'MsgBox Prop.Name
        Select Case Prop.Name
            Case "Cliente_PN"
                Prop.Value = Range("clientPN").Value
            Case "IMX_PN"
                Prop.Value = Range("interplexPn").Value
            Case "Cliente"
                Prop.Value = Range("client").Value
            Case "Fecha_Creado"
                Prop.Value = Range("date").Value
            Case "Costo_Total"
                Prop.Value = Range("total").Value
            Case "Problema"
                Prop.Value = Range("problem").Value
            Case "Originador"
                Prop.Value = Range("originator").Value
            Case "Disposicion"
                Prop.Value = Range("disposition").Value
            Case "Inspeccion"
                Prop.Value = Range("inspection").Value
            Case "Cargos_Cliente"
                Prop.Value = Range("charges").Value
            Case "CAR"
                Prop.Value = Range("car").Value
            Case "#RMA"
                Prop.Value = Range("rmaName").Value
            Case "#ID"
                Prop.Value = Str(Format(Range("number"), "000#"))
            Case "Comentarios"
                Prop.Value = Range("comments").Value
            Case "Numero_ID"
                Prop.Value = Str(Format(Range("number"), "000#"))
            Case "Scrap"
                Prop.Value = Range("scrap").Value
            Case "PO_Return"
                Prop.Value = Range("poreturn").Value
            Case "Cantidad_Piezas"
                Prop.Value = Range("quantity").Value
            Case "Costo_Unitario"
                Prop.Value = Range("unitPrice").Value
            Case Else
                'N/A
        End Select
    Next Prop
    
End Sub