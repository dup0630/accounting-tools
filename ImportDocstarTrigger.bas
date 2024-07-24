Sub ImportDocstarGuillevin()
    Dim table_name As String
    Dim worksheet_name As String
    table_name = "DCSTR"
    worksheet_name = "Docstar Guillevin"
    Module8.ClearGuillevin
    ImportDocstarTool table_name, worksheet_name
End Sub
Sub ImportDocstarBrogan()
    Dim table_name As String
    Dim worksheet_name As String
    table_name = "DCSTRBRGN"
    worksheet_name = "Docstar Brogan"
    Module8.ClearBrogan
    ImportDocstarTool table_name, worksheet_name
End Sub
Sub ImportDocstarDubo()
    Dim table_name As String
    Dim worksheet_name As String
    table_name = "DCSTRDUBO"
    worksheet_name = "Docstar Dubo"
    Module8.ClearDubo
    ImportDocstarTool table_name, worksheet_name
End Sub