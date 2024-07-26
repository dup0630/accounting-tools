Sub ImportWorkday()
    Dim xlsxFilePath As String
    Dim csvFilePath As String
    Dim wb As Workbook
    Dim qt As QueryTable
    Dim ws As Worksheet

    Set ws = ThisWorkbook.Sheets("Workday")
    If ws.Range("A2") <> "" Then
        response = MsgBox ("This action will overwrite the previous data. Are you sure you want to continue?", vbYesNo + vbQuestion, "Guillevin International Inc.")
        If response = vbNo Then
            Exit Sub
        End If
    End If

    ' CONVERT TO CSV
    xlsxFilePath = GetUserSelectedFile("Please select Workday data file:")
    If xlsxFilePath = "" Then
        MsgBox "No file selected."
        Exit Sub
    End If
    csvFilePath = Left(xlsxFilePath, Len(xlsxFilePath) - 4) & "csv"

    Set wb = Workbooks.Open(xlsxFilePath)
    wb.Sheets(1).SaveAs Filename:=csvFilePath, FileFormat:=xlCSV
    wb.Close SaveChanges:=False
    
    ' IMPORT DATA
    ws.Cells.Delete
    Set qt = ws.QueryTables.Add(Connection:="TEXT;" & csvFilePath, Destination:=ws.range("A1"))
    
    With qt
        .TextFileParseType = xlDelimited
        .TextFileCommaDelimiter = True
        .TextFileColumnDataTypes = Array(1)
        .Refresh BackgroundQuery:=False
    End With
    qt.Delete
    Kill csvFilePath
    
    ' CONVERT DATA INTO TABLE
    Dim tbl As ListObject
    Set tbl = ws.ListObjects.Add(SourceType:=xlSrcRange, _
                                 Source:=ws.range("A1").CurrentRegion, _
                                 xlListObjectHasHeaders:=xlYes)
    tbl.Name = "WD"
    tbl.Range.Columns.AutoFit

    ' DELETE UNNECESSARY COLUMNS (DIFFERENT FOR EACH DATABASE!)
    ws.Columns("A:E").Delete Shift:=xlToLeft
    
    ' CHANGE INVOICE NO INTO NUMBER TYPE WITH TEXT TO COLUMNS
    ws.range("A:A").TextToColumns Destination:=range("A:A"), _
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
    
    Sheets("Statement").ListObjects("TABLE").ListColumns("Workday Status").DataBodyRange.Formula = "=VLOOKUP([@[Inv. number]],WD,7,FALSE)"
    ' Sheets("Statement").range("E2").FormulaR1C1 = "=VLOOKUP(RC[-4],WD,7,FALSE)"
End Sub
