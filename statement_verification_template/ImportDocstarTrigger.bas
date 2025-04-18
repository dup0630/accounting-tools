Sub NewDocstar()
    Dim new_ws As Worksheet
    Dim helper_ws As Worksheet
    Dim table_name As String
    Dim worksheet_name As String
    Dim n As Integer
    
    Application.CutCopyMode = False ' Clear clipboard
    
    n = ThisWorkbook.Sheets("Config").Range("B3").Value
    
    table_name = "DCSTR" & (n + 1)
    worksheet_name = "Docstar" & (n + 1)
    
    On Error Resume Next
    Set new_ws = ThisWorkbook.Sheets(worksheet_name)
    On Error GoTo 0
    ' Checks if ws is Docstar1
    If new_ws Is Nothing Then
        Set new_ws = ThisWorkbook.Worksheets.Add(After:=Worksheets(Worksheets.Count))
        new_ws.Name = worksheet_name
        new_ws.Tab.Color = RGB(0, 32, 96)
        Sheet1.Activate
    End If
    
    ImportDocstarTool table_name, worksheet_name
    ThisWorkbook.Sheets("Config").Range("B3").Value = ThisWorkbook.Sheets("Config").Range("B3").Value + 1
    Application.CutCopyMode = False ' Clear clipboard
End Sub