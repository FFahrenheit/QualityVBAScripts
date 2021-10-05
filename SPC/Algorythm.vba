'' This module loads the data into the inspection list with an
'' linear search algorythm

Private Type Lengths
    Data As Integer
    Dict As Integer
    EData As Integer
    EDict As Integer
End Type

Public Sub GetData()
    Dim Data As Worksheet
    Dim Dict As Worksheet
    Dim I, J As Integer
    Dim L As Lengths
    Set Data = ActiveWorkbook.Worksheets("Data")
    Set Dict = ActiveWorkbook.Worksheets("Diccionario")
    
    L = FullDict(Data, Dict)
    
    If L.EData <> L.EDict Then
        MsgBox "Solo se usarán los valores que tengan referencia. Para asegurar " _
        & "el llenado automatico, por favor complete el diccionario", vbCritical + vbOKOnly, _
        "Diccionario incompleto"
    End If
    
    I = 5
    
    While I <= L.Data
        
    Wend

End Sub

Function LastCell(SheetName As String)
    
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

