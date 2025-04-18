Sub ImportDocstarTool(nm As String, wsnm As String)
    Dim filePath As String
    Dim csvFilePath As String
    Dim qt As QueryTable
    Dim colName As String
    Dim colIndex As Long
    Dim colNamesToDelete As Collection
    Dim DocstarColumns As Variant
    Dim tbl As ListObject
    Dim dcstr_lang As String
    Dim ws As Worksheet
    Dim amount_col As Range
    Dim wb As Workbook


    ' CHECK DOCSTAR SELECTED LANGUAGE
    dcstr_lang = Sheet1.ComboBox1.Value

    ' IMPORT FILE
    Set ws = ThisWorkbook.Sheets(wsnm)
    ws.Cells.Delete
    filePath = GetUserSelectedFile("Please select Docstar data file:")
    If filePath = "" Then
        MsgBox "No file selected."
        If ws.Name <> "Docstar1" Then
            Application.DisplayAlerts = False
            ws.Delete
            Application.DisplayAlerts = True
        End If
        
        n = ThisWorkbook.Sheets("Config").Range("B3").Value
        ThisWorkbook.Sheets("Config").Range("B3").Value = n - 1
        
        Exit Sub
    End If

    If Right$(filePath, 5) = ".xlsx" Then
        DocstarColumns = Array("PC", "WorkflowStep", "InvoiceNumber")
        ' CREATE CSV COPY
        csvFilePath = Left(filePath, Len(filePath) - 4) & "csv"
        Set wb = Workbooks.Open(filePath)
        ' REMOVE UNNECESSARY DATA
        wb.Sheets(1).Rows("1:3").Delete Shift:=xlUp
        wb.Sheets(1).SaveAs Filename:=csvFilePath, FileFormat:=xlCSV

        wb.Close SaveChanges:=False

        ' IMPORT CSV
        ws.Cells.Delete
        Set qt = ws.QueryTables.Add(Connection:="TEXT;" & csvFilePath, Destination:=ws.Range("A1"))
        With qt
            .TextFileParseType = xlDelimited
            .TextFileCommaDelimiter = True
            .TextFileColumnDataTypes = Array(1)
            .Refresh BackgroundQuery:=False
        End With
        qt.Delete
        Kill csvFilePath
    ElseIf Right$(filePath, 4) = ".csv" Then
        ' LANGUAGE SETTINGS
        If dcstr_lang = "ENGLISH" Then
            DocstarColumns = Array("Branch", "Workflow Step", "InvoiceNum")
        ElseIf dcstr_lang = "FRANÇAIS" Then
            DocstarColumns = Array("Branch", "Ã‰tape du flux de travail", "InvoiceNum")
        Else
            ' Handle unexpected value
            MsgBox "Please select a language for Docstar.", vbCritical, "Error"
            Exit Sub
        End If

        ' IMPORT CSV
        ws.Cells.Delete
        Set qt = ws.QueryTables.Add(Connection:="TEXT;" & filePath, Destination:=ws.Range("A1"))
        With qt
            .TextFileParseType = xlDelimited
            .TextFileCommaDelimiter = True
            .TextFileColumnDataTypes = Array(1)
            .Refresh BackgroundQuery:=False
        End With
        qt.Delete
    End If
    
    ' CREATE TABLE
    Set tbl = ws.ListObjects.Add(SourceType:=xlSrcRange, _
                                 Source:=ws.Range("A1").CurrentRegion, _
                                 xlListObjectHasHeaders:=xlYes)
    
    tbl.Name = nm
    ' REARRANGE COLUMNS 2
    On Error GoTo LangErrorHandler
    For i = LBound(DocstarColumns) To UBound(DocstarColumns)
        colName = DocstarColumns(i)
        colIndex = tbl.ListColumns(colName).Index
        If colIndex > 1 Then
            tbl.ListColumns(colName).Range.Cut
            tbl.ListColumns(1).Range.Insert Shift:=xlToRight
            Application.CutCopyMode = False
        End If
    Next i
    tbl.Range.Columns.AutoFit
    
    ' CHANGE INVOICE# INTO NUMBER TYPE WITH TEXT TO COLUMNS
    ws.Range("A:A").TextToColumns Destination:=Range("A:A"), _
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
        

    ' ENTER FORMULAS
    Sheet1.ListObjects("TABLE").ListColumns("Docstar WF Step").DataBodyRange.Formula = "=VLOOKUP([@[Inv. number]]," & tbl.Name & ", 2, FALSE)"
    Sheet1.ListObjects("TABLE").ListColumns("Branch").DataBodyRange.Formula = "=VLOOKUP([@[Inv. number]]," & tbl.Name & ",3,FALSE)"
    
    If Sheets("Config").Range("B4").Value = False Then
        Sheets("Config").Range("B4").Value = True
    End If
    
    Exit Sub
    
LangErrorHandler:
    MsgBox "Error: Wrong language selected. Please try again.", vbExclamation
    ws.Cells.Delete
    Exit Sub
End Sub

