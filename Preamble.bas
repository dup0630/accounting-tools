Public Function GetUserSelectedFile(ByVal promptMessage As String) As String
    Dim filePath As Variant
    filePath = Application.GetOpenFilename("All Files,*.*", , promptMessage)

    If filePath <> False Then
        GetUserSelectedFile = filePath
    Else
        GetUserSelectedFile = ""
    End If
End Function
Public Function ExtractNumber(cell_value As String)
    For i = Len(cell_value) To 1 Step -1
        current_char = Mid(cell_value, i, 1)
        If Len(output) = 2 And current_char = "," Then
            output = "." & output
        ElseIf IsNumeric(current_char) Or current_char = "." Then
            output = current_char & output
        End If
    Next i

    ExtractNumber = output
End Function