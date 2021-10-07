Sub SaveRMA()
    On Error GoTo Fault
    Result = Validate()
    
    If Result = False Then
        'MsgBox "No valido!"

        Exit Sub
    End If
    
    'Validación de demás campos
    
    Message = "El RMA será guardado en el SharePoint de QMS en la libreria de IMX_RMA.  " _
    & " Aqui podra consultarla para posterior revisión y/o edición." _
    & "¿Desea continuar?"
    
    Style = vbYesNo + vbExclamation + vbDefaultButton2
    
    Response = MsgBox(Message, Style, "SALVAR ARCHIVO")
    
    If Response = vbYes Then
    
        Filename = "RMA-" & Format(Range("number"), "000#")
        
        ThisWorkbook.IsSaved = True
        
        ActiveWorkbook.SaveAs Filename:= _
        "https://interplexgroup.sharepoint.com/americas/imx/imx_qms/IMX_RMA/" & Filename & ".xlsm" _
        , FileFormat:=xlOpenXMLWorkbookMacroEnabled, Password:="", WriteResPassword:="", _
        ReadOnlyRecommended:=False, CreateBackup:=False
        
        ActiveSheet.Shapes.Range(Array("uploadButton")).Select
        Selection.Delete
        Range("A1").Select
        
        Msg = "El RMA ha sido guardado con éxito"
        Style = vbOKOnly + vbInformation
        
        MsgBox Msg, Style, "RMA GUARDADO"
        
    End If    
    Exit Sub
    
Fault:
    Msg = "El RMA no puede ser grabado, por favor" _
        & " asegurese que no haber utilizado caracteres especiales en el nombre como " _
        & "(!?¡*[]&$()%@/), esto evita que el DMR pueda ser guardada, " _
        & "los guiones bajos ( _ ) SI pueden ser utilizados" & Chr(10) & Chr(10) & "En caso contrario por favor " _
        & "informe al administrador del sharepoint sobre el problema"
    Style = vbOKOnly + vbInformation
    
    MsgBox Msg, Style, "ERROR AL GUARDAR"
    MsgBox "Error: " & Err.Description, vbOKOnly + vbExclamation, "ERROR"
End Sub

