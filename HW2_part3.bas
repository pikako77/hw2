Attribute VB_Name = "part3"
Sub StockVolTotal()
Attribute StockVolTotal.VB_ProcData.VB_Invoke_Func = "t\n14"
   Dim ABSENT As Integer 'define absent value as constant
   ABSENT = -9999
   
   Dim ws As Worksheet
   
   Dim ticker As String
   Dim total As Double
   Dim lastRow As Long         ' last row index
   Dim count_ticker As Integer ' Use to write output in the worksheet
   Dim total_ticker As Integer ' Use to scan the percent change
   
   Dim Price_opening As Double         ' first price of the year at the opening
   Dim Price_closing As Double         ' last price of the year at the closing
   Dim yearly_change As Double         ' store yearly change
   Dim yearly_change_percent As Double ' store yearly change in %
   
   Dim min_index As Long            ' store the index (i.e line number) of the min
   Dim min_tmp As Double            ' store temporary min
   
   Dim max_index As Long            ' store the index (i.e line number) of the max
   Dim max_tmp As Double            ' store temporary max increase
   
   Dim max_vol_index As Long        ' store the index (i.e line number) of the max volume
   Dim max_vol_tmp As Double        ' store temporary max volume
   

   For Each ws In Worksheets
    ' Initialize parameters for min and max search
      min_index = 1                    ' initialize to 1
      min_tmp = Abs(ABSENT)            ' set min big enough to not miss the true min
      
      max_vol_index = 1                ' initialize to 1
      max_vol_tmp = ABSENT             ' set max small enough to not miss the true max

      max_index = 1                    ' initialize to 1
      max_tmp = 2 * ABSENT               ' set max small enough to not miss the true max

   
    ' Write the header for the additional colums
     ws.Range("I1").Value = "Ticker"
     ws.Range("J1").Value = "Yearly change"
     ws.Range("K1").Value = "Percent change"
     ws.Range("L1").Value = "Total stock volume"
     ws.Range("A1:P1").Font.Bold = True  ' put all the header in bold (easier to read)
   
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
             
    total_ticker = count_ticker ' rename the total ticker
     
    For j = 2 To total_ticker + 1 ' last row for ticker = total_ticker+1
     
        ' find max increase
        If ((ws.Cells(j, 11).Value <> "NaN") And (ws.Cells(j, 11).Value < min_tmp)) Then ' find min value
            min_tmp = ws.Cells(j, 11).Value
            min_index = j
        End If
        
        ' find max decrease
       If ((ws.Cells(j, 11).Value <> "NaN") And ws.Cells(j, 11).Value > max_tmp) Then   ' find min value
            max_tmp = ws.Cells(j, 11).Value
            max_index = j
        End If
        
        ' find max volume
        If ((ws.Cells(j, 11).Value <> "NaN") And ws.Cells(j, 12).Value > max_vol_tmp) Then  ' find min value
            max_vol_tmp = ws.Cells(j, 12).Value
            max_vol_index = j
        End If
    Next j
     
    ws.Cells(1, 15).Value = "Ticker"
    ws.Cells(1, 16).Value = "Value"
    
    ws.Cells(2, 14).Value = "Greatest % Increase"
    ws.Cells(2, 15).Value = ws.Cells(max_index, 9).Value
    ws.Cells(2, 16).Value = ws.Cells(max_index, 11).Value
    ws.Cells(2, 16).NumberFormat = "0.00%" ' Format
    'ws.Cells(2, 17).Value = max_index
     
    ws.Cells(3, 14).Value = "Greatest % Decrease"
    ws.Cells(3, 15).Value = ws.Cells(min_index, 9).Value
    ws.Cells(3, 16).Value = ws.Cells(min_index, 11).Value
    ws.Cells(3, 16).NumberFormat = "0.00%" ' Format
    ' ws.Cells(3, 17).Value = min_index
     
    ws.Cells(4, 14).Value = "Greatest total volume "
    ws.Cells(4, 15).Value = ws.Cells(max_vol_index, 9).Value
    ws.Cells(4, 16).Value = ws.Cells(max_vol_index, 12).Value
    'ws.Cells(4, 17).Value = max_vol_index

    ' Define column width
    ws.Range("N:N").ColumnWidth = 20
    ws.Range("O:O").ColumnWidth = 6
    ws.Range("P:P").ColumnWidth = 15
    
    ws.Range("N2:N4").Font.Bold = True
   Next ws

End Sub


