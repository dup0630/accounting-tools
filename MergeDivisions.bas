Sub MergeGuillevinBrogan()
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim SheetNames(1 To 3) As String
    Dim TableNames(1 To 3) As String
    Dim TableExists(1 To 3) As Boolean

    SheetNames(1) = "Docstar Guillevin"
    SheetNames(2) = "Docstar Brogan"
    SheetNames(3) = "Docstar Dubo"
    TableNames(1) = "DCSTR"
    TableNames(2) = "DCSTRBRGN"
    TableNames(3) = "DCSTRDUBO"
    TableExists(1) = False
    TableExists(2) = False
    TableExists(3) = False

    For i = 1 To 3
        Set ws = ThisWorkbook.Worksheets(SheetNames(i))
        For Each tbl In ws.ListObjects
            If tbl.Name = TableNames(i) Then
                TableExists(i) = True
            End If
        Next tbl
    Next i

    For j = 1 To 3
        If Not TableExists(j) Then
            Set ws = ThisWorkbook.Worksheets(SheetNames(j))
            Set tbl = ws.ListObjects.Add(SourceType:=xlSrcRange, _
                                            Source:=ws.range("A1:A2"), _
                                            xlListObjectHasHeaders:=xlNo)
            tbl.Name = TableNames(j)
        End If
    Next j

    ' ENTER FORMULAS
    Sheet1.ListObjects("TABLE").ListColumns("Docstar WF Step").DataBodyRange.Formula = "=IFNA(VLOOKUP([@[Inv. number]],DCSTR, 2, FALSE), IFNA(VLOOKUP([@[Inv. number]],DCSTRBRGN, 2, FALSE), VLOOKUP([@[Inv. number]],DCSTRDUBO, 2, FALSE)))"
    ' Sheets("Statement").ListObjects("TABLE").ListColumns("Amount match (Y/N)").DataBodyRange.Formula = "=IFNA(IF([@Amount]=VLOOKUP([@[Inv. number]], DCSTR,3,FALSE),""Y"",""N""), IFNA(IF([@Amount]=VLOOKUP([@[Inv. number]], DCSTRBRGN,3,FALSE),""Y"",""N""), IF([@Amount]=VLOOKUP([@[Inv. number]], DCSTRDUBO,3,FALSE),""Y"",""N"")))"

End Sub