Private Sub Workbook_Open()
    TemplateName = "IXFC055_Formato_DMR.xlsm"
    If ThisWorkbook.Name = TemplateName Then
        Application.AutoRecover.Enabled = False
        ActiveWorkbook.AutoSaveOn = False
    End If
End Sub

Private Sub Workbook_BeforeSave(ByVal SaveAsUI As Boolean, Cancel As Boolean)
    Call SetServerProperties
    TemplateName = "IXFC055_Formato_DMR.xlsm"
    If ThisWorkbook.Name = TemplateName Then
    
        Response = MsgBox("Está a punto de hacer cambios en el FORMATO. ¿Está seguro" _
        & " de esto? Recuerde que para guardar un nuevo DMR debe de presionar el" _
        & " botón GUARDAR DMR EN SP", vbYesNoCancel + vbCritical, "GUARDAR EN PLATILLA")
        
        If Response = vbYes Then
            MsgBox "Formato de DMR actualizado", vbOKOnly, "PLANTILLA GUARDADA"
        Else
            Cancel = True
            MsgBox "No se han guardado los cambios en la plantilla", vbOKOnly, "PLANTILLA NO ACTUALIZADA"
        End If
        
    End If
End Sub