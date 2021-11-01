Sub IdentifyFormat()
    Dim Filename As String
    Filename = GetFile()
    If Filename <> "" Then
        Formato = Identify(Filename)
        
        MsgBox "Formato: " & Formato
    End If
End Sub

Function Identify(Filename As String) As String
    On Error GoTo ErrHandler
    
    Dim Source As Worksheet
    Dim src As Workbook
    Application.ScreenUpdating = False
    'MsgBox "Filename: " & Filename
    
    Set src = Workbooks.Open(Filename, True, True)
    
    Y = src.Name
    
    MsgBox Y
    Set Source = Workbooks(Y).Sheets(1)
    
    'Identify format "B", all different
    
    I = 1
    If Source.Range("A1") <> "" Then
        I = I + 1
        Do While Source.Range("A" & I) <> "" And IsNumeric(Source.Range("A" & I)) = True
            I = I + 1
        Loop
        MsgBox "Registros = " & I - 1
    End If
    
    If I > 3 Then
        Identify = "B"
        Exit Function
    End If
    
    'Identify format "A"
        
    I = 10
    OneMeasure = True
    Count = 0
    
    MsgBox "Formato A"
    
    Do While Source.Range("B" & I) <> ""
        J = I + 2
        
        Do While Source.Range("G" & J) <> ""
        
            If IsNumeric(Source.Range("G" & J)) = False Then    'Formato desconocido
                GoTo NoResult
            End If
            
            J = J + 1
        Loop
        
        Count = Count + 1
        I = J + 1
    Loop
    
    MsgBox "Registros = " & Count
    
    If Count > 1 Then
        If OneMeasure = True Then
            Identify = "A1"
        Else
            Identify = "A2"
        End If
        Exit Function
    Else
        GoTo NoResult
    End If
    
        
    'Close all
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    src.Close (False)
    Set src = Nothing
    Exit Function
ErrHandler:
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    src.Close (False)
    Set src = Nothing
    MsgBox "Error: " & Err.Description
    MsgBox "No se pudieron cargar los datos de la hoja de inspeccion", vbOKOnly + vbCritical, "Error de carga"
NoResult:
    Identify = ""
    Exit Function
End Function
