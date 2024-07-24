Sub ClearWD()
    Dim ws As Worksheet
    Dim response
    Set ws = ThisWorkbook.Sheets("Workday")
    response = MsgBox ("Do you want to clear the data in " & ws.Name & "?", vbYesNo + vbQuestion, "Guillevin International Inc.")
    If response = vbYes Then
        ws.Cells.Delete
        ThisWorkbook.Sheets("Statement").ListObjects("TABLE").ListColumns("Docstar WF Step").DataBodyRange.ClearContents
        ThisWorkbook.Sheets("Statement").ListObjects("TABLE").ListColumns("Amount match (Y/N)").DataBodyRange.ClearContents
    Else
        Exit Sub
    End If
End Sub
Sub ClearGuillevin()
    Dim ws As Worksheet
    Dim response
    Set ws = ThisWorkbook.Sheets("Docstar Guillevin")
    response = MsgBox ("Do you want to clear the data in " & ws.Name & "?", vbYesNo + vbQuestion, "Guillevin International Inc.")
    If response = vbYes Then
        ws.Cells.Delete
        ThisWorkbook.Sheets("Statement").ListObjects("TABLE").ListColumns("Docstar WF Step").DataBodyRange.ClearContents
        ThisWorkbook.Sheets("Statement").ListObjects("TABLE").ListColumns("Amount match (Y/N)").DataBodyRange.ClearContents
    Else
        Exit Sub
    End If
End Sub
Sub ClearBrogan()
    Dim ws As Worksheet
    Dim response
    Set ws = ThisWorkbook.Sheets("Docstar Brogan")
    response = MsgBox ("Do you want to clear the data in " & ws.Name & "?", vbYesNo + vbQuestion, "Guillevin International Inc.")
    If response = vbYes Then
        ws.Cells.Delete
        ThisWorkbook.Sheets("Statement").ListObjects("TABLE").ListColumns("Docstar WF Step").DataBodyRange.ClearContents
        ThisWorkbook.Sheets("Statement").ListObjects("TABLE").ListColumns("Amount match (Y/N)").DataBodyRange.ClearContents
    Else
        Exit Sub
    End If
End Sub
Sub ClearDubo()
    Dim ws As Worksheet
    Dim response
    Set ws = ThisWorkbook.Sheets("Docstar Dubo")
    response = MsgBox ("Do you want to clear the data in " & ws.Name & "?", vbYesNo + vbQuestion, "Guillevin International Inc.")
    If response = vbYes Then
        ws.Cells.Delete
        ThisWorkbook.Sheets("Statement").ListObjects("TABLE").ListColumns("Docstar WF Step").DataBodyRange.ClearContents
        ThisWorkbook.Sheets("Statement").ListObjects("TABLE").ListColumns("Amount match (Y/N)").DataBodyRange.ClearContents
    Else
        Exit Sub
    End If
End Sub