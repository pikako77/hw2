Attribute VB_Name = "part2"
Sub StockVolTotal()
Attribute StockVolTotal.VB_ProcData.VB_Invoke_Func = "t\n14"
   Dim ws As Worksheet
   
   Dim ticker As String
   Dim total As Double
   Dim lastRow As Long         ' last row index
   Dim count_ticker As Integer ' Use to write output in the worksheet
   
   Dim Price_opening As Double         ' first price of the year at the opening
   Dim Price_closing As Double         ' last price of the year at the closing
   Dim yearly_change As Double         ' store yearly change
   Dim yearly_change_percent As Double ' store yearly change in %
   
   Dim ABSENT As Integer 'define absent value as constant
   
   ABSENT = -9999
   
   For Each ws In Worksheets
   
    ' Write the header for the additional colums
     ws.Range("I1").Value = "Ticker"
     ws.Range("J1").Value = "Yearly change"
     ws.Range("K1").Value = "Percent change"
     ws.Range("L1").Value = "Total stock volume"
     ws.Range("A1:L1").Font.Bold = True  ' put all the header in bold (easier to read)
   
     ' get the last data row number (i.e get number of data in the sheet)
     lastRow = ws.Cells(Rows.Count, "A").End(xlUp).Row

     ' initialize total and ticker counter
     total = 0
     count_ticker = 0
     
     'set first opening price
     Price_opening = ws.Cells(2, 3)
     
     For i = 2 To lastRow
        
        total = total + ws.Cells(i, 7).Value
        
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
           ticker = ws.Cells(i, 1).Value   ' get ticker
           count_ticker = count_ticker + 1 ' increase the ticker total by one
           
           Price_closing = ws.Cells(i, 6)                         ' get the last price of the year at the closing
           yearly_change = Price_closing - Price_opening          ' Compute the yearly change
                     
           ' Use for debug
           'ws.Cells(count_ticker + 1, 14).Value = Price_opening
           'ws.Cells(count_ticker + 1, 15).Value = Price_closing
           
           If (Price_opening = 0) Then
              yearly_change_percent = ABSENT
              
              ' Use for debug
              'ws.Cells(count_ticker + 1, 16).Value = i
           Else
              yearly_change_percent = yearly_change / Price_opening  ' Compute the yearly change in percent
           End If
           
           
           ' Write in worksheet
           ws.Cells(count_ticker + 1, 9).Value = ticker
           ws.Cells(count_ticker + 1, 10).Value = yearly_change
           ws.Cells(count_ticker + 1, 10).NumberFormat = "0.000000000"
           If (yearly_change_percent = ABSENT) Then
              ws.Cells(count_ticker + 1, 11).Value = "NaN"
              ws.Cells(count_ticker + 1, 11).Font.Color = vbRed
              ws.Cells(count_ticker + 1, 11).Font.Bold = True
              ws.Cells(count_ticker + 1, 11).HorizontalAlignment = xlCenter
           Else
              ws.Cells(count_ticker + 1, 11).Value = yearly_change_percent
           End If
            
           ws.Cells(count_ticker + 1, 12).Value = total
           ws.Cells(count_ticker + 1, 11).NumberFormat = "0.00%"
           
           
           
           ' Color the yearly_change backgroud
           If (yearly_change < 0) Then
              ws.Cells(count_ticker + 1, 10).Interior.ColorIndex = 3 ' red = 3
           Else
              ws.Cells(count_ticker + 1, 10).Interior.ColorIndex = 43 'green = 4 is too shiny for me... use 43 instead
           End If
           
           ' reset total for the next ticker
           total = 0
           
           ' get opening price for the next ticker
           Price_opening = ws.Cells(i + 1, 3)
 
        End If
     Next i
        ' Define column width
           ws.Range("I:I").ColumnWidth = 6
           ws.Range("J:L").ColumnWidth = 16
           
     
   Next ws

End Sub
