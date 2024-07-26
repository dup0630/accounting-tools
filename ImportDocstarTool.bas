Sub ImportDocstarTool(nm As String, wsnm As String)
    Dim csvFilePath As String
    Dim qt As QueryTable
    Dim colName As String
    Dim colIndex As Long
    Dim colNamesToDelete As Collection
    Dim DocstarColumns As Variant
    Dim tbl As ListObject
    Dim dcstr_lang As String
    dcstr_lang = ThisWorkbook.Sheets("Statement").ComboBox1.Value
    Dim ws As Worksheet
    Dim amount_col As Range


    ' CHECK DOCSTAR SELECTED LANGUAGE
    If dcstr_lang = "ENGLISH" Then
        DocstarColumns = Array("InvoiceNum", "InvoiceAmt", "PONum", "Workflow Step", "InvoiceDate", "Branch", "Annotation Text")
    ElseIf dcstr_lang = "FRANÇAIS" Then
        DocstarColumns = Array("InvoiceNum", "InvoiceAmt", "PONum", "Ã‰tape du flux de travail", "InvoiceDate", "Branch", "Texte de l'annotation")
    Else
        ' Handle unexpected value
        MsgBox "Please select a language for Docstar.", vbCritical, "Error"
        Exit Sub
    End If

    ' IMPORT CSV
    Set ws = ThisWorkbook.Sheets(wsnm)
    ws.Cells.Delete
    csvFilePath = GetUserSelectedFile("Please select Docstar data file:")
    If csvFilePath = "" Then
        MsgBox "No file selected."
        Exit Sub
    End If
    Set qt = ws.QueryTables.Add(Connection:="TEXT;" & csvFilePath, Destination:=ws.range("A1"))
    
    With qt
        .TextFileParseType = xlDelimited
        .TextFileCommaDelimiter = True
        .TextFileColumnDataTypes = Array(1)
        .Refresh BackgroundQuery:=False
    End With
    qt.Delete
    
    ' CREATE TABLE
    Set tbl = ws.ListObjects.Add(SourceType:=xlSrcRange, _
                                 Source:=ws.range("A1").CurrentRegion, _
                                 xlListObjectHasHeaders:=xlYes)
    
    tbl.Name = nm

     ' REARRANGE COLUMNS
    On Error GoTo LangErrorHandler
    For i = LBound(DocstarColumns) To UBound(DocstarColumns)
        colName = DocstarColumns(i)
        colIndex = tbl.ListColumns(colName).Index
        
        ' Move the column to the desired position
        If colIndex <> i + 1 Then
            tbl.ListColumns(colName).range.Cut
            tbl.ListColumns(i + 1).range.Insert Shift:=xlToRight
        End If
    Next i
    
    
    
    ' DELETE UNNECESSARY COLUMNS
    Set colNamesToDelete = New Collection
    
    For Each col In tbl.ListColumns
        colName = col.Name
        If IsError(Application.Match(colName, DocstarColumns, 0)) Then
            colNamesToDelete.Add colName
        End If
    Next col
    For i = colNamesToDelete.Count To 1 Step -1
        tbl.ListColumns(colNamesToDelete(i)).Delete
    Next i
    tbl.range.Columns.AutoFit
    
    ' CHANGE INVOICE# INTO NUMBER TYPE WITH TEXT TO COLUMNS
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

    ' CHANCE AMOUNT INTO NUMBERS ONLY(REMOVE $)
    Set amount_col = tbl.ListColumns("InvoiceAmt").DataBodyRange
    For Each cell In amount_col
        If cell.Value <> "" Then cell.Value = ExtractNumber(cell.Value)
    Next cell

    ' ENTER FORMULAS
    Sheets("Statement").ListObjects("TABLE").ListColumns("Docstar WF Step").DataBodyRange.Formula = "=VLOOKUP([@[Inv. number]]," & tbl.Name & ", 4, FALSE)"
    ' Sheets("Statement").range("F2").FormulaR1C1 = "=VLOOKUP(RC[-5]," & tbl.Name & ", 4, FALSE)"
    
    Sheets("Statement").ListObjects("TABLE").ListColumns("Amount match (Y/N)").DataBodyRange.Formula = "=IF([@Amount]=VLOOKUP([@[Inv. number]], " & tbl.Name & ",2,FALSE),""Y"",""N"")"
    ' Sheets("Statement").range("G2").FormulaR1C1 = "=IF(RC[-3]=VLOOKUP(RC[-6], " & tbl.Name & ",2,FALSE),""Y"",""N"")"
    
    Exit Sub
    
LangErrorHandler:
    MsgBox "Error: Wrong language selected. Please try again.", vbExclamation
    ws.Cells.Delete
    Exit Sub
End Sub



