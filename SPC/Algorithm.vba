'' This module loads the data into the inspection list with an
'' linear search algorythm

Private Type Lengths
    Data As Integer  'Last row in Data
    Dict As Integer  'Last row in Dict
    EData As Integer 'Size of Data
    EDict As Integer 'Size of Dict
End Type

Public Sub GetData()
    Dim Property As String
    
    Dim Data As Worksheet
    Dim Dict As Worksheet
    Dim Insp As Worksheet
    Dim InspName As String
    
    Dim I As Integer
    Dim L As Lengths
    InspName = "Hoja de inspeccion"
    
    Set Data = ActiveWorkbook.Worksheets("Data")
    Set Dict = ActiveWorkbook.Worksheets("Diccionario")
    Set Insp = ActiveWorkbook.Worksheets(InspName)
    
    L = FullDict(Data, Dict)
    
    If L.EData <> L.EDict Then
        MsgBox "Solo se usarán los valores que tengan referencia. Para asegurar " _
        & "el llenado automatico, por favor complete el diccionario", vbCritical + vbOKOnly, _
        "Diccionario incompleto"
    End If
    
    C = LastCell(InspName)
    Success = True
    
    I = 5
    
    While I <= L.Data
        Property = Data.Range("A" & I).Value
        Value = Data.Range("B" & I).Value
        
        Destination = SearchValue(Dict, L, Property)
        If Destination <> 0 Then
            Insp.Cells(Destination, C) = Value
        Else
            Success = False
        End If
        I = I + 1
    Wend
    
    If Success = True Then
        MsgBox "Se han importa los datos de forma correcta", _
        vbOKOnly + vbInformation, "Hoja importada"
    Else
        MsgBox "Se han importado parcialmente los datos, llene las celdas faltantes", _
        vbOKOnly + vbCritical, "Hoja importada"
    End If
    

End Sub

Function LastCell(SheetName As String)
    Dim I As Integer
    Dim WS As Worksheet
    Set WS = ActiveWorkbook.Sheets(SheetName)
    I = 19
    
    Do While True
        If ((WS.Cells(14, I).Value = "-" Or WS.Cells(14, I) = "") And WS.Cells(21, I).Value = "" _
            And WS.Cells(22, I).Value = "") Or I > 60 Then
            Exit Do
        End If
        I = I + 1
    Loop
    
    LastCell = I
    MsgBox "Last cell = " & LastCell
    
End Function

Private Function FullDict(Data As Worksheet, Dict As Worksheet) As Lengths
    Dim FunLengths As Lengths
    Dim EntriesData As Integer
    Dim EntriesDict As Integer
    
    With Data
        LastRowData = .Cells(.Rows.Count, "B").End(xlUp).Row
        EntriesData = Application.WorksheetFunction.CountA(.Range("B5:B" & LastRowData))
    End With
    ' MsgBox EntriesData
    
    With Dict
        LastRowDict = .Cells(.Rows.Count, "C").End(xlUp).Row
        EntriesDict = Application.WorksheetFunction.CountA(.Range("C5:C" & LastRowDict))
    End With
    
    With FullDict
        .Data = LastRowData
        .Dict = LastRowDict
        .EData = EntriesData
        .EDict = EntriesDict
    End With
    
End Function

Function SearchValue(Dict As Worksheet, L As Lengths, Property As String)
    I = 5
    SearchValue = 0
    
    While I <= L.Dict
        If Dict.Range("B" & I) = Property Then
            SearchValue = Dict.Range("C" & I)
            Exit Function
        End If
        
        I = I + 1
    Wend
End Function

