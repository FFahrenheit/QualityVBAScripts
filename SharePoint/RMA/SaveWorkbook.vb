Public IsSaved As Boolean

Private Sub Workbook_Open()
    TemplateName = "IXFC256_Formato_RMA.xlsm"
    If ThisWorkbook.Name = TemplateName Then
        Application.AutoRecover.Enabled = False
        ActiveWorkbook.AutoSaveOn = False
    End If
    
    ThisWorkbook.IsSaved = False
End Sub

Private Sub Workbook_BeforeSave(ByVal SaveAsUI As Boolean, Cancel As Boolean)

    TemplateName = "IXFC256_Formato_RMA.xlsm"
    If ThisWorkbook.Name = TemplateName And ThisWorkbook.IsSaved = False Then
        'Cancel = True
        Response = MsgBox("Está a punto de hacer cambios en el FORMATO. ¿Está seguro" _
        & " de esto? Recuerde que para guardar un nuevo RMA debe de presionar el" _
        & " botón GUARDAR DMR EN SP", vbYesNoCancel + vbQuestion, "GUARDAR EN PLATILLA")
        
        If Response = vbYes Then
            MsgBox "Formato de RMA actualizado", vbOKOnly, "PLANTILLA GUARDADA"
        Else
            Cancel = True
            MsgBox "No se han guardado los cambios en la plantilla", vbOKOnly, "PLANTILLA NO ACTUALIZADA"
        End If
        
    End If
    
    Call SetServerProperties
    ThisWorkbook.IsSaved = False
End Sub

