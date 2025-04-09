Sub InsertColumn()
    Dim tbl As ListObject
    Dim colIndex As Integer
    
    Set tbl = Sheet1.ListObjects("TABLE")
    colIndex = tbl.ListColumns("Workday Status").DataBodyRange.Column
    
    Sheet1.Columns(colIndex).Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
End Sub
