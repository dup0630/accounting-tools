Sub MergeDocstar()
    Dim ws As Worksheet
    Dim new_ws As Worksheet
    Dim tbl As ListObject
    Dim new_tbl As ListObject
    Dim paste_row As Long
    Dim SheetNames() As String
    Dim TableNames() As String
    Dim n As Integer
    
    n = ThisWorkbook.Sheets("Config").Range("B3").Value
    If n = 0 Then
        MsgBox "Data has not been imported yet.", vbExclamation, "Guillevin International Inc."
        Exit Sub
    End If
    
    ReDim SheetNames(1 To n)
    ReDim TableNames(1 To n)
    
    For i = 1 To n
        SheetNames(i) = "Docstar" & i
        TableNames(i) = "DCSTR" & i
    Next i
    
    Set new_ws = ThisWorkbook.Worksheets.Add(After:=Worksheets(Worksheets.Count))
    On Error GoTo NameTaken
    new_ws.Name = "MergedDocstarData"
    On Error GoTo 0
    new_ws.Tab.Color = RGB(0, 32, 96)
    
    ' MERGE DATA
    paste_row = 1
    For i = 1 To n
        On Error Resume Next
        Set ws = ThisWorkbook.Sheets(SheetNames(i))
        Set tbl = ws.ListObjects(TableNames(i))
        If i = 1 Then
            tbl.HeaderRowRange.Copy
            new_ws.Cells(paste_row, 1).PasteSpecial Paste:=xlPasteValues
            Application.CutCopyMode = False ' Clear clipboard
            paste_row = paste_row + 1
        End If
        
        tbl.DataBodyRange.Copy
        new_ws.Cells(paste_row, 1).PasteSpecial Paste:=xlPasteValues
        paste_row = new_ws.Cells(new_ws.Rows.Count, 1).End(xlUp).Row + 1
        Application.CutCopyMode = False ' Clear clipboard
    Next i
    On Error GoTo 0
    new_ws.Columns.AutoFit
    
    ' CREATE TABLE
    Set new_tbl = new_ws.ListObjects.Add(SourceType:=xlSrcRange, _
                                         Source:=new_ws.Range("A1").CurrentRegion, _
                                         xlListObjectHasHeaders:=xlYes)
    
    new_tbl.Name = "DCSTRMERGE"
    Sheets("MergedDocstarData").Range("A1").Select
    
    ' ENTER FORMULAS
    Sheet1.ListObjects("TABLE").ListColumns("Docstar WF Step").DataBodyRange.Formula = "=VLOOKUP([@[Inv. number]]," & new_tbl.Name & ", 2, FALSE)"
    Sheet1.ListObjects("TABLE").ListColumns("Branch").DataBodyRange.Formula = "=VLOOKUP([@[Inv. number]]," & new_tbl.Name & ",3,FALSE)"
    
    Sheet1.Activate
    Application.CutCopyMode = False ' Clear clipboard
    MsgBox "Merge completed.", vbInformation, "Guillevin International Inc."
    
    Exit Sub
    
NameTaken:
    MsgBox "Error: Data has already been merged. Please click on 'Clear Merge' and try again.", vbExclamation
    Application.DisplayAlerts = False
    new_ws.Delete
    Application.DisplayAlerts = True
    Sheet1.Activate
    Exit Sub
    
End Sub
