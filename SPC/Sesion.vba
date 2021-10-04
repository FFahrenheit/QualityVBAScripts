Sub CerrarSesion()
    On Error GoTo Handler
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    ActiveSheet.Shapes("Log in").Visible = True
    ActiveSheet.Shapes("Log out").Visible = False
    ActiveSheet.Shapes("User icon").Visible = False
    
    Worksheets("Hoja de inspeccion").Shapes("Generar hoja").Visible = False
    Worksheets("Hoja de inspeccion").Shapes("generateTemplate").Visible = False
    Worksheets("Hoja de inspeccion").Shapes("Zoom+").Visible = False
    Worksheets("Hoja de inspeccion").Shapes("Zoom-").Visible = False
    Worksheets("Hoja de inspeccion").Shapes("CC").Visible = False
    
    Changecolor
    ActiveSheet.Shapes("Graphic 17").Visible = False
    Application.ScreenUpdating = False
    Sheets("Hoja de inspeccion").Select
    ActiveSheet.Unprotect ("Calidad2020")
    
    Range("E5,E8,E11,L11,L8,L5,S11,S8,S5,W8,W11,B17:G1301,I22:P1301,b1304:l1306").Select
    Selection.Locked = True
    
    Worksheets("Usuarios").Cells(7, 8) = 5
    
    For Each Worksheet In ActiveWorkbook.Worksheets
        Worksheet.Protect ("Calidad2020")
        'Worksheet.Visible = True
    Next Worksheet

    Sheets(Array("HOME", "Hoja de inspeccion", "SPC", "Nuevo analisis", _
        "PLAN DE ACCION NUEVO", "Usuarios", "Correo")).Select
    Sheets("HOME").Activate
    ActiveWindow.DisplayHeadings = False
    Sheets("HOME").Select
    
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    ActiveWindow.DisplayWorkbookTabs = False
    ' Application.ExecuteExcel4Macro "SHOW.TOOLBAR(""Ribbon"",False)"
    Application.DisplayFormulaBar = False
    Application.DisplayStatusBar = False
Handler:
    Resume Next
End Sub

Sub IniciarSesion()
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Usuarios.Show
    
    If Worksheets("Usuarios").Cells(7, 8).Value <> 5 Then
        ActiveSheet.Shapes("Log in").Visible = False
        ActiveSheet.Shapes("Log out").Visible = True
        Worksheets("Hoja de inspeccion").Shapes("Generar hoja").Visible = True
        Worksheets("Hoja de inspeccion").Shapes("generateTemplate").Visible = True
        Worksheets("Diccionario").Unprotect ("Calidad2020")
        Worksheets("Hoja de inspeccion").Shapes("Zoom+").Visible = True
        Worksheets("Hoja de inspeccion").Shapes("Zoom-").Visible = True
        Worksheets("Hoja de inspeccion").Shapes("CC").Visible = True
        ActiveSheet.Shapes("Graphic 17").Visible = True

        'Desbloquear celdas de hoja de inspeccion

        Sheets("Hoja de inspeccion").Select
        ActiveSheet.Unprotect ("Calidad2020")
        Range("E5,E8,E11,L11,L8,L5,S5,S8,S11,W8,W11,B17:G1301,I22:P1301,b1304:l1306").Select
        Selection.Locked = False
  
        Range("E5").Select
        Sheets("Hoja de inspeccion").Protect ("Calidad2020"), DrawingObjects:=False
        Sheets("HOME").Select
    End If

    'Privilegios de usuario

    Select Case Worksheets("Usuarios").Cells(7, 8).Value ' Entrada de caso
        Case 0 'Master
            'Application.ExecuteExcel4Macro "SHOW.TOOLBAR(""Ribbon"",True)"
            Application.DisplayFormulaBar = True
            Application.DisplayStatusBar = True
            
            For Each Worksheet In ActiveWorkbook.Worksheets
                Worksheet.Unprotect ("Calidad2020")
                Worksheet.Visible = True
            Next Worksheet
            
            Worksheets("PLAN DE ACCION").Visible = False
            Worksheets("Nombres").Visible = False
            Worksheets("Analisis").Visible = False
        
            Sheets(Array("HOME", "Hoja de inspeccion", "SPC", "Nuevo analisis", _
            "PLAN DE ACCION NUEVO", "Usuarios", "Correo")).Select
            Sheets("HOME").Activate
            ActiveWindow.DisplayHeadings = True

            ActiveWindow.DisplayWorkbookTabs = True

            Sheets("HOME").Select
        
        Case 3   ' Administrador
            ActiveSheet.Shapes("User icon").Visible = True
    End Select
    
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
End Sub

