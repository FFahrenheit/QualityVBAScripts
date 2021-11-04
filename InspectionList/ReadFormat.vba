Sub ReadFormat()
    Dim Filename As String
    Filename = GetFile()
    If Filename <> "" Then
        Formato = Identify(Filename)
        MsgBox "Formato: " & Formato
        
        Select Case Formato
            Case "A1"
                CopyTemplateA1 (Filename)
            Case "B"
                CopyTemplateB (Filename)
            Case "A2"
                CopyTemplateA3 (Filename)
            Case "A3"
                CopyTemplateA3 (Filename)
            Case Else
                MsgBox "Formato no soportado"
                Exit Sub
        End Select
        
        MsgBox "Formato pasado a diccionario"
    End If
End Sub
