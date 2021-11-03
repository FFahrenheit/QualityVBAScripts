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
    
    Set src = Workbooks.Open(Filename, True, True)
    
    Y = src.Name
    
    Set Source = Workbooks(Y).Sheets(1)
    
    'Identify format "B", all different
    
    I = 1
    If Source.Range("A1") <> "" Then
        I = I + 1
        
        Do While Source.Range("A" & I) <> "" And IsNumeric(Source.Range("A" & I)) = True
            I = I + 1
        Loop
        J = 2 'Cotas
        I = I - 2 'Mediciones
        
        Do While Source.Cells(I, J) <> "" And IsNumeric(Source.Cells(I, J)) = True
            J = J + 1
        Loop
        
        J = J - 1
        
        MsgBox "Cotas = " & J & ", Mediciones = " & I
    End If
    
    If I > 3 And J > 1 Then
        Identify = "B"
        GoTo Return0
    End If
    
    'Identify format "A"
    
    
    I = 10
    MedicionesMax = 0
    Count = 0
    
    Do While Source.Range("B" & I) <> ""
    
        J = I + 2
        
        Mediciones = 0
        Do While Source.Range("B" & J) <> ""
            
            If (IsNumeric(Source.Range("H" & J)) = False And _
                IsNumeric(Source.Range("G" & J)) = False) Then   'Formato desconocido
                GoTo NoResult
            End If
            Mediciones = Mediciones + 1
            J = J + 1
        Loop
        
        ' Mediciones = Mediciones - 1
        
        If Mediciones > MedicionesMax Then MedicionesMax = Mediciones
        
        Count = Count + 1
        I = J + 1
    Loop
    
    'MsgBox "Mediciones Max = " & MedicionesMax
    
    If Count > 1 Then
        MsgBox "Cotas = " & Count & " Maximo mediciones: " & MedicionesMax
        If MedicionesMax = 1 Then
            Identify = "A1"
        ElseIf MedicionesMax = 2 Then
            Identify = "A2"
        ElseIf MedicionesMax > 2 Then
            Identify = "A3"
        End If
        
        GoTo Return0
        
    Else
        GoTo NoResult
    End If
    
NoResult:
    Identify = ""
Return0:
    'Close all
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    src.Close (False)
    Set src = Nothing
    Exit Function
ErrHandler:
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    'src.Close (False)
    'Set src = Nothing
    MsgBox "Error: " & Err.Description
    MsgBox "No se pudieron cargar los datos de la hoja de inspeccion", vbOKOnly + vbCritical, "Error de carga"
End Function
