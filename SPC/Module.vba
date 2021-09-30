'' This module was coded by -i.lopez-
'  any concerns or questions please address to i.lopez@mx.interplex.com
'  This module contains the functions to auto load measures into the
'  the inspection list automatically
'  v.0.0.1

'' This functions opens a file explorer dialog and
'' return the string of the path to the file (excel files)
Function GetFile() As String
    Dim SeletedFile As String
    SelectedFile = ""

    With Application.FileDialog(msoFileDialogFilePicker)
        .ButtonName = "Confirm"
        .AllowMultiSelect = False
        .Title = "Seleccionar el archivo generado"
        .Filters.Clear
        .Filters.Add "Excel Worksheets", "*.xls; *.xlsx; *.xlsm"
        .InitialFileName = "D:\"
        If .Show = -1 Then
        'ok
            SelectedFile = .SelectedItems(1)
            GetFile = SelectedFile
            ' OpenWorkbook (SelectedFile)
        Else
        'cancel
            MsgBox "No se pudieron cargar los datos de la hoja de inspeccion", vbOKOnly + vbCritical, "Error de carga"
        End If
        
    End With
End Function

'' This subroutine gets the template file and calls the function to process it
Sub LoadTemplate()
    Message = "Esta opcion le permitirá cargar la plantilla de lectura de las hojas de inspeccion, " _
    & "usela solo si desea modificar la plantilla. ¿Desea continuar?"
    Response = MsgBox(Message, vbYesNo + vbQuestion + vbDefaultButton1, "CAMBIAR PLANTLLA")
    
    If Response = vbYes Then
        Filename = GetFile()
        If Filename <> "" Then
            MsgBox "Hola"
            'CopyTemplate (Filename)
        End If
    End If
End Sub

