Sub SaveAlert()
    On Error GoTo Fault
    Result = Validate()
    'MsgBox "Result = " & Result
    
    If Result = False Then
        Exit Sub
    End If
    
    Message = "El DMR será guardado en el SharePoint de QMS en la libreria de IMX_DMR.  " _
        & " Aqui podra consultarla para posterior revisión y/o edición." _
        & "¿Desea continuar?"
    
    Style = vbYesNo + vbExclamation + vbDefaultButton2
    
    Response = MsgBox(Message, Style, "SALVAR ARCHIVO")
    
    If Response = vbYes Then
    
        Filename = "DMR-" & Format(Range("number"), "000#")
        
        ThisWorkbook.IsSaved = True
        
        ActiveWorkbook.SaveAs Filename:= _
        "https://interplexgroup.sharepoint.com/americas/imx/imx_qms/IMX_DMR/" & Filename & ".xlsm" _
        , FileFormat:=xlOpenXMLWorkbookMacroEnabled, Password:="", WriteResPassword:="", _
        ReadOnlyRecommended:=False, CreateBackup:=False
        
        Worksheets("DMR Hoja 1").Activate
        ActiveSheet.Shapes.Range(Array("uploadButton")).Select
        Selection.Delete
        Range("A1").Select
        
        Worksheets("DMR Hoja 2").Activate
        ActiveSheet.Shapes.Range(Array("uploadButton2")).Select
        Selection.Delete
        'ActiveSheet.Shapes.Range(Array("newAction")).Select
        'Selection.Delete
        'Range("A1").Select
        
        Worksheets("DMR Hoja 1").Activate
        
        Msg = "El DMR ha sido guardado con éxito"
        Style = vbOKOnly + vbInformation
        
        MsgBox Msg, Style, "DMR GUARDADO"
        
    End If
    
    Exit Sub

Fault:
    Msg = "El DMR no puede ser grabado, por favor" _
        & " asegurese que no haber utilizado caracteres especiales en el nombre como " _
        & "(!?¡*[]&$()%@/), esto evita que el DMR pueda ser guardada, " _
        & "los guiones bajos ( _ ) SI pueden ser utilizados" & Chr(10) & Chr(10) & "En caso contrario por favor " _
        & "informe al administrador del sharepoint sobre el problema"
    Style = vbOKOnly + vbInformation
    
    MsgBox Msg, Style, "ERROR AL GUARDAR"
End Sub