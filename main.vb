Function parseValue(value As String) As String
    If IsEmpty(value) Then
        parseValue = ""
    Else
        parseValue = value
    End If
End Function

Sub convertJson()
    
    Dim c           As Collection
    Dim d           As Dictionary        'Add reference to scripting runtime
    Dim v           As Dictionary
    Dim json        As String
    
    Set c = New Collection
    Set d = New Dictionary
    
    d.Add "ExcelTest", c
    For Each cell in Range("K18")        'Adapt as you need
        Set v = New Dictionary
        v.Add Range("J4").value, parseValue(cell.Offset(0,0).value)
        v.Add Range("L4").value, parseValue(cell.Offset(0,1).value)
        c.Add v
    Next
    
    json = jsonConverter.ConvertToJson(d)
    
    Debug.Print json
End Sub
