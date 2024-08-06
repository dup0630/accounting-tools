Sub FixColumns()
    Dim ws As Worksheet
    Dim config As Worksheet
    Dim tbl As ListObject
    Dim rng As Range
    Dim amount As Range
    Dim cols() As Range

    Set ws = Sheet1
    Set tbl = ws.ListObjects("TABLE")
    Set amount = tbl.ListColumns("Amount").DataBodyRange
    ReDim cols(1 to tbl.ListColumns.Count)
    For i = 1 To tbl.ListColumns.Count 
        If tbl.ListColumns(i).Name = "Workday Status" Then
            Exit For
        End If
        Set rng = tbl.ListColumns(i).DataBodyRange
        If Application.WorksheetFunction.CountA(rng) > 0 Then
            rng.TextToColumns Destination:=rng, _
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
        End If
    Next i
    
    For Each cell In amount
        If cell.Value <> "" Then cell.Value = ExtractNumber(cell.Value)
    Next cell

    MsgBox "Columns have been formatted."
End Sub