Sub ClearWD()
    Dim ws As Worksheet
    Dim response
    
    WorkdayContainsData = ThisWorkbook.Sheets("Config").Range("B5").Value
    If WorkdayContainsData = False Then
        MsgBox "Nothing to clear.", vbExclamation, "Guillevin International Inc."
        Exit Sub
    End If
    
    Set ws = ThisWorkbook.Sheets("Workday")
    response = MsgBox("Do you want to clear the data in " & ws.Name & "?", vbYesNo + vbQuestion, "Guillevin International Inc.")
    If response = vbYes Then
        ws.Cells.Delete
        Sheet1.ListObjects("TABLE").ListColumns("Workday Status").DataBodyRange.ClearContents
        Sheet1.ListObjects("TABLE").ListColumns("Workday Amount").DataBodyRange.ClearContents
        Sheet1.ListObjects("TABLE").ListColumns("Payment Date").DataBodyRange.ClearContents
        ThisWorkbook.Sheets("Config").Range("B5").Value = False
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
        Sheet1.ListObjects("TABLE").ListColumns("Branch").DataBodyRange.ClearContents
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
        Sheet1.ListObjects("TABLE").ListColumns("Branch").DataBodyRange.ClearContents
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
        Sheet1.ListObjects("TABLE").ListColumns("Branch").DataBodyRange.ClearContents
    Else
        Exit Sub
    End If
End Sub
Sub ClearDS()
    Dim ws As Worksheet
    Dim response
    Dim n As Integer
    Dim DocstarContainsData As Boolean
    DocstarContainsData = ThisWorkbook.Sheets("Config").Range("B4").Value
    If DocstarContainsData = False Then
        MsgBox "Nothing to clear.", vbExclamation, "Guillevin International Inc."
        Exit Sub
    End If
    response = MsgBox("Do you want to clear the data in all Docstar worksheets?", vbYesNo + vbQuestion, "Guillevin International Inc.")
    If response = vbYes Then
        ThisWorkbook.Sheets("Docstar1").Cells.Delete
        n = ThisWorkbook.Sheets("Config").Range("B3").Value
        Application.DisplayAlerts = False  ' Disable alerts
        On Error Resume Next ' ----------------------
        For i = 2 To n
            Set ws = ThisWorkbook.Sheets("Docstar" & i)
            If Not ws Is Nothing Then
                ws.Delete
                n = n - 1
            End If
        Next i
        On Error GoTo 0 ' ---------------------------
        Application.DisplayAlerts = True  ' Re-enable alerts
        Sheet1.ListObjects("TABLE").ListColumns("Docstar WF Step").DataBodyRange.ClearContents
        Sheet1.ListObjects("TABLE").ListColumns("Branch").DataBodyRange.ClearContents
        ThisWorkbook.Sheets("Config").Range("B3").Value = n - 1
        ThisWorkbook.Sheets("Config").Range("B4").Value = False
    Else
        Exit Sub
    End If
    
End Sub
Sub ClearMerge()
    Dim ws As Worksheet
    Dim response
    
    On Error GoTo NoMerge
    Set ws = ThisWorkbook.Sheets("MergedDocstarData")
    On Error GoTo 0
    response = MsgBox("Do you want to clear the data in " & ws.Name & "? The remaining Docstar sheets will not be affected", vbYesNo + vbQuestion, "Guillevin International Inc.")
    If response = vbYes Then
        Application.DisplayAlerts = False
        ws.Delete
        Application.DisplayAlerts = True
        Sheet1.ListObjects("TABLE").ListColumns("Docstar WF Step").DataBodyRange.ClearContents
        Sheet1.ListObjects("TABLE").ListColumns("Branch").DataBodyRange.ClearContents
    Else
        Exit Sub
    End If
    Exit Sub
    
NoMerge:
    MsgBox "Error: Data has not been merged. Nothing to clear.", vbExclamation
    Exit Sub
    
End Sub
