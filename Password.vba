Private Sub CommandButton1_Click()
Dim pword As String pword = Application.InputBox("Enter password", "Password Required", Type:=2) 
If pword <> "password" Then Exit SubActiveWorkbook.Sheets("INSTRUCTIONS & SQL").Visible = xlSheetVisible
ActiveWorkbook.Sheets("Backlog Workorders").Visible = xlSheetVisible
ActiveWorkbook.Sheets("Current Workorders").Visible = xlSheetVisible
End Sub
