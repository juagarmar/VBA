Sub test2()
Dim LastRow As Long
Dim i As Long
Dim TargetRange As Range
On Error Resume Next
    LastRow = Sheets("DATA").Cells(Rows.Count, "A").End(xlUp).Row
    Set TargetRange = Sheets("CONSOLE").Range("A2:C" & LastRow)
    MsgBox LastRow
For i = 2 To LastRow
'i = 4
    Sheets("DATA").Cells(i, "B").Value = Application.WorksheetFunction.VLookup(Sheets("DATA").Cells(i, "A"), TargetRange, 2, False)
    Sheets("DATA").Cells(i, "C").Value = Application.WorksheetFunction.VLookup(Sheets("DATA").Cells(i, "A"), TargetRange, 3, False)
Next i
Sheets("CONSOLE").Cells.Clear
End Sub
