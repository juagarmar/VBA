Sub deletebydateinarow()
Dim d1 As DateDim
lngCount As LongDim
i As Longd1 = DateSerial(Year(Now), Month(Now), 0)
lngCount = Application.WorksheetFunction.CountA(Columns(1))
For i = 2 To lngCount
If Sheets("Sheet1").Cells(i, 3) > d1 Then
Sheets("Sheet1").Cells(i, 3).EntireRow.Delete
End If
Next i
End Sub
