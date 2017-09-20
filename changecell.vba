Private Sub Worksheet_Change(ByVal Target As Range)
Dim KeyCells As RangeSet KeyCells = Range("G8")    
If Not Application.Intersect(KeyCells, Range(Target.Address)) _           
Is Nothing 
Then Call new5           
End Sub
