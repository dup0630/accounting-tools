Sub InsertColumn()
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim col As Range
    Dim colIndex As Integer
    
    Set ws = ThisWorkbook.Worksheets("Statement")
    Set tbl = ws.ListObjects("TABLE")
    colIndex = tbl.ListColumns("Workday Status").DataBodyRange.Column
    
    ws.Columns(colIndex).Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
End Sub
