Sub ClearWD()
    Dim ws As Worksheet
    Dim response
    Set ws = ThisWorkbook.Sheets("Workday")
    response = MsgBox("Do you want to clear the data in " & ws.Name & "?", vbYesNo + vbQuestion, "Guillevin International Inc.")
    If response = vbYes Then
        ws.Cells.Delete
        Sheet1.ListObjects("TABLE").ListColumns("Workday Status").DataBodyRange.ClearContents
        Sheet1.ListObjects("TABLE").ListColumns("Workday Amount").DataBodyRange.ClearContents
        Sheet1.ListObjects("TABLE").ListColumns("Payment Date").DataBodyRange.ClearContents
    Else
        Exit Sub
    End If
End Sub
Sub ClearGuillevin()
    Dim ws As Worksheet
    Dim response
    Set ws = ThisWorkbook.Sheets("Docstar Guillevin")
    response = MsgBox("Do you want to clear the data in " & ws.Name & "?", vbYesNo + vbQuestion, "Guillevin International Inc.")
    If response = vbYes Then
        ws.Cells.Delete
        Sheet1.ListObjects("TABLE").ListColumns("Docstar WF Step").DataBodyRange.ClearContents
        Sheet1.ListObjects("TABLE").ListColumns("Amount match (Y/N)").DataBodyRange.ClearContents
    Else
        Exit Sub
    End If
End Sub
Sub ClearBrogan()
    Dim ws As Worksheet
    Dim response
    Set ws = ThisWorkbook.Sheets("Docstar Brogan")
    response = MsgBox("Do you want to clear the data in " & ws.Name & "?", vbYesNo + vbQuestion, "Guillevin International Inc.")
    If response = vbYes Then
        ws.Cells.Delete
        Sheet1.ListObjects("TABLE").ListColumns("Docstar WF Step").DataBodyRange.ClearContents
        Sheet1.ListObjects("TABLE").ListColumns("Amount match (Y/N)").DataBodyRange.ClearContents
    Else
        Exit Sub
    End If
End Sub
Sub ClearDubo()
    Dim ws As Worksheet
    Dim response
    Set ws = ThisWorkbook.Sheets("Docstar Dubo")
    response = MsgBox("Do you want to clear the data in " & ws.Name & "?", vbYesNo + vbQuestion, "Guillevin International Inc.")
    If response = vbYes Then
        ws.Cells.Delete
        Sheet1.ListObjects("TABLE").ListColumns("Docstar WF Step").DataBodyRange.ClearContents
        Sheet1.ListObjects("TABLE").ListColumns("Amount match (Y/N)").DataBodyRange.ClearContents
    Else
        Exit Sub
    End If
End Sub
