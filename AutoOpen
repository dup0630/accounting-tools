Sub Auto_Open()
    Dim newFileName As String
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Config")
    If ws.range("A1").Value = False Then
        newFileName = Application.GetSaveAsFilename(FileFilter:="Excel Macro-Enabled Workbook (*.xlsm), *.xlsm", _
                                                     Title:="Enter New File Name", _
                                                     InitialFileName:="newfile.xlsm")
        If newFileName <> "False" Then
            ThisWorkbook.SaveAs Filename:=newFileName, FileFormat:=xlOpenXMLWorkbookMacroEnabled
        End If
        ws.range("A1").Value = True
    Else
        Exit Sub
    End If
End Sub
