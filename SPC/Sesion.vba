Sub IniciarSesion()
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    Usuarios.Show

    If Worksheets("Usuarios").Cells(7, 8).Value <> 5 Then

            ActiveSheet.Shapes("Log in").Visible = False
            ActiveSheet.Shapes("Log out").Visible = True
            Worksheets("Hoja de inspeccion").Shapes("Generar hoja").Visible = True
            Worksheets("Hoja de inspeccion").Shapes("generateTemplate").Visible = True
    'Mas texto'
    End If
End Sub

Sub CerrarSesion()
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
        ActiveSheet.Shapes("Log in").Visible = True
        ActiveSheet.Shapes("Log out").Visible = False
        ActiveSheet.Shapes("User icon").Visible = False
        Worksheets("Hoja de inspeccion").Shapes("Generar hoja").Visible = False
        Worksheets("Hoja de inspeccion").Shapes("generateTemplate").Visible = False
        'Mas texto'

End Sub
