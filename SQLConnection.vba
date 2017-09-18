Sub SQLconnection()
Dim conn As ADODB.ConnectionDim rs As ADODB.RecordsetDim sConnString As StringDim SQLKPI1 As String
SQLKPI1 = Sheets("SQL").Range("C2").Value         
' Create the connection string.    
sConnString = "Provider=SQLOLEDB.1;Password=xxxxx;Persist Security Info=True;User ID=CMMS_EAMDATA_JCI;
Initial Catalog=xxxxxP;Data Source=xxxxxx;Use Procedure for Prepare=1;
Auto Translate=True;Packet Size=4096;Workstation ID=RTPWL11H25966;
Use Encryption for Data=False;Tag with column collation when possible=False"        
' Create the Connection and Recordset objects.    
Set conn = New ADODB.Connection    
Set rs = New ADODB.Recordset        
' Open the connection and execute.    
conn.Open sConnString    
Set rs = conn.Execute(SQLKPI1)        
' Check we have data.    
If Not rs.EOF Then            
' Transfer result.        
Sheets("KPI 01 Data").Range("A" & Rows.Count).End(xlUp)(2).CopyFromRecordset rs            
' Close the recordset        
rs.Close    
Else        
MsgBox "Error: No records returned.", 
vbCritical    
End If
    ' Clean up    
If CBool(conn.State And adStateOpen) Then conn.Close    
Set conn = Nothing    
Set rs = Nothing   
End Sub
