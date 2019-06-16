Attribute VB_Name = "Module11"
Sub StockVoltotal()
Attribute StockVoltotal.VB_ProcData.VB_Invoke_Func = "t\n14"
   Dim ticker As String
   Dim total As Double  'it can be long integer too
   
   Dim lastRow As Long ' last row index
   Dim count_ticker As Integer
   Dim ws As Worksheet
   

   
   For Each ws In Worksheets
   
    ' Write the header for the additional colums
     ws.Range("I1").Value = "Ticker"
     ws.Range("J1").Value = "Total stock volume"
     ws.Range("A1:J1").Font.Bold = True
   
     lastRow = ws.Cells(Rows.Count, "A").End(xlUp).Row
     'lastRow = Cells(Rows.Count, "A").End(xlUp).Row  ' for 1 sheet
     'MsgBox (lastRow)
       
     total = 0  ' initialize total
     count_ticker = 0
     
     'MsgBox (total)
     For i = 2 To lastRow
        
        total = total + Cells(i, 7).Value
        
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
           ticker = ws.Cells(i, 1).Value
           count_ticker = count_ticker + 1
           ws.Cells(count_ticker + 1, 9).Value = ticker
           ws.Cells(count_ticker + 1, 10).Value = (total)
           
           total = 0 ' reset total for the next ticker
        End If
               
     Next i
     
   Next ws

End Sub
