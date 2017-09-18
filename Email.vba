Sub sendemail()
Dim OApp As Object, OMail As Object, signature As StringDim iCounter As Integer   
Set OApp = CreateObject("Outlook.Application")
Set OMail = OApp.CreateItem(0)
Set sh = Sheets("INSTRUCTIONS & SQL")
With OMail    
.Display    
End With        
signature = OMail.Body
With OMail    
.To = sh.Range("B7")    
.CC = sh.Range("B8")    
.BCC = sh.Range("B9")    
.Subject = "Daily Work Order Report    " & Format(Date, "DD-MMM-YYYY")    
.Body = sh.Range("B11").Value & vbNewLine & sh.Range("B12").Value & vbNewLine & sh.Range("B13").Value & vbNewLine & sh.Range("B14").Value & vbNewLine & sh.Range("B15").Value & vbNewLine & sh.Range("B16").Value & vbNewLine & sh.Range("B12").Value & signature    
.Attachments.Add ActiveWorkbook.FullName    
.SentOnBehalfOfName = "dataanalysis"    
End With    
Set OMail = NothingSet OApp = Nothing
End Sub
