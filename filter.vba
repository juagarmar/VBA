Sub filterpm()Dim 
sh11 As WorksheetDim 
sh22 As WorksheetSet 
sh11 = Sheets("Revlist")Set 
sh22 = Sheets("RevData")
sh22.SelectRange("A6:AT6").SelectSelection.AutoFilterSelection.AutoFilter Field:=26, Criteria1:=sh11.Range("B2").Value', Operator:=xlOr, Criteria2:=""'Array( _'        "   IL", "   IN", "   MI", "   OH", "   WV"), Operator:=xlFilterValuesCall critfield1Call critfield2End Sub
