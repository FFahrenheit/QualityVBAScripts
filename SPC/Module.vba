'' This module was coded by -i.lopez-
'  any concerns or questions please address to i.lopez@mx.interplex.com
'  This module contains the functions to auto load measures into the
'  the inspection list automatically
'  v.0.0.1


'' This subroutine gets the template file and calls the function to process it
Sub LoadTemplate()
    Message = "Esta opcion le permitirá cargar la plantilla de lectura de las hojas de inspeccion, " _
    & "usela solo si desea modificar la plantilla. ¿Desea continuar?"
    Response = MsgBox(Message, vbYesNo + vbQuestion + vbDefaultButton1, "CAMBIAR PLANTLLA")
    
    If Response = vbYes Then
        Filename = Utils.GetFile()
        If Filename <> "" Then
            CopyTemplate (Filename)
        End If
    End If
End Sub

'This function opens and loads the template
Sub CopyTemplate(Filename As String)
    Dim A As String
    Dim C As String
    
    A = ActiveWorkbook.Name  'Workbook destino
    C = "Diccionario"
    
    Exists = Utils.SheetExists(C, A)
    If Exists = False Then
        'No existe, la creamos
        With Workbooks(A)
            Set ws = .Sheets.Add(After:=.Sheets(.Sheets.Count))
            ws.Name = C
        End With
    End If
    
    Dim Destino As Worksheet
    Dim Dict As Worksheet
    Dim src As Workbook
    
    Application.ScreenUpdating = False
    'MsgBox "Filename: " & Filename
    
    Set src = Workbooks.Open(Filename, True, True)
    
    Y = src.Name
    Z = "Sheet1"
    
    Set Destino = Workbooks(Y).Worksheets(Z)
    Set Dict = Workbooks(A).Worksheets(C)
    Dict.Columns(2).ClearContents
    
    'Header details
    Dict.Range("B1") = "Pieza"
    Dict.Range("B2") = "Fecha"
    Dict.Range("B3") = "Hora"

    Dict.Range("C1") = "N/A"
    Dict.Range("C2") = 14
    Dict.Range("C3") = 21
    
    i = 10
    j = 5
    
    While Destino.Range("B" & i) <> ""
        Cota = Destino.Range("B" & i)
        Dict.Range("B" & j) = Cota
        
        i = i + 4
        j = j + 1
    Wend
    
    'Close all
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    src.Close (False)
    Set src = Nothing
    
    'Show tabs
    If Sheets("Usuarios").ProtectContents = True Then
        Sheets("Hoja de inspeccion").Unprotect ("Calidad2020")
        Sheets("Diccionario").Unprotect ("Calidad2020")
        Sheets("Hoja de inspeccion").Visible = True
        Sheets("SPC").Visible = True
        Sheets("PLAN DE ACCION NUEVO").Visible = True
        Sheets("HOME").Visible = True
        Sheets("Nuevo analisis").Visible = False

        Sheets("Usuarios").Visible = False
        Sheets("Correo").Visible = False
        Sheets("Nombres").Visible = False
        Sheets("Analisis").Visible = False
        Sheets("PLAN DE ACCION").Visible = False
    End If
    
    Application.DisplayFormulaBar = True
    Application.DisplayStatusBar = True

    Sheets("Hoja de inspeccion").Unprotect ("Calidad2020")
    Sheets("Diccionario").Unprotect ("Calidad2020")
    Sheets("Hoja de inspeccion").Visible = True
    Sheets("Diccionario").Visible = True
    ActiveWindow.DisplayHeadings = True
    ActiveWindow.DisplayWorkbookTabs = True
    Dict.Activate
    Application.DisplayFormulaBar = True
    Exit Sub
ErrHandler:
    MsgBox "Error: " & Err.Description
    MsgBox "No se pudo cargar la plantilla de la hoja de inspeccion", vbOKOnly + vbCritical, "Error de carga"
    Exit Sub
End Sub

