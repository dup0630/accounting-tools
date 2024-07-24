Sub FixColumns()
    Dim stat As Worksheet
    Dim config As Worksheet
    Dim tbl As ListObject
    Dim rng As Range
    Dim amount As String

    Set stat = ThisWorkbook.Sheets("Statement")
    Set tbl = stat.ListObjects("TABLE")
    Set amount = tbl.ListColumns("Amount").DataBodyRange
    cols = Array("A", "B", "C", "D")
    
    For i = 0 To 3
        rng = cols(i) & ":" & cols(i)
        Sheets("Statement").range(rng).TextToColumns Destination:=Sheets("Statement").range(rng), _
            DataType:=xlDelimited, _
            TextQualifier:=xlDoubleQuote, _
            ConsecutiveDelimiter:=False, _
            Tab:=False, _
            Semicolon:=False, _
            Comma:=False, _
            Space:=False, _
            Other:=False, _
            FieldInfo:=Array(1, 1), _
            TrailingMinusNumbers:=True
    Next i
    
    For Each cell In amount
        If cell.Value <> "" Then cell.Value = ExtractNumber(cell.Value)
    Next cell

    MsgBox "Columns have been formatted."
End Sub