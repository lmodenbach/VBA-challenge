'Assignment: create 4 new columns: unique ticker symbol, yearly change in stock,
'yearly change as a percentage, total volume of each ticker symbol
'Bonus: create a table showing top percent increase, bottom percent decrease, and largest total volume

'new macro
Sub stocks_challenge()

'loop through worksheets
  For Each ws In Worksheets

'declare variables
    Dim ticker_symbol As String
    Dim year_opening_price, year_ending_price, yearly_change, percent_yearly_change As Double
    Dim total_volume, new_table_row, greatest_total_volume, lastRowK, lastRowL, target_row As Long
    Dim greatest_percent_increase, greatest_percent_decrease As Double
      
'initialize starting row of new table
    new_table_row = 2

'loop through the rows
    For Row = 2 To (ws.Cells(Rows.Count, 1).End(xlUp).Row)

'if it's the start of a series grab year_opening_price and opening volume
      If ws.Cells(Row, 1).Value <> ws.Cells(Row - 1, 1).Value Then
        year_opening_price = ws.Cells(Row, 3).Value
        total_volume = ws.Cells(Row, 7).Value

'afternote: this task could have been done outside of for loop (but within for each loop)
'insert column headers
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
         
'if it's the end of a series record ticker symbol and total volume, calculate yearly change
      ElseIf ws.Cells(Row, 1).Value <> ws.Cells(Row + 1, 1).Value Then
        total_volume = total_volume + ws.Cells(Row, 7).Value
        ticker_symbol = ws.Cells(Row, 1)
        ws.Cells(new_table_row, 9).Value = ticker_symbol
        ws.Cells(new_table_row, 12).Value = total_volume
        year_ending_price = ws.Cells(Row, 6).Value
        yearly_change = year_ending_price - year_opening_price
        ws.Cells(new_table_row, 10).Value = yearly_change

'color code yearly change (red = -, green = +)
        If yearly_change < 0 Then
          ws.Cells(new_table_row, 10).Interior.ColorIndex = 3
          ws.Cells(new_table_row, 10).Font.ColorIndex = 1
        Else
          ws.Cells(new_table_row, 10).Interior.ColorIndex = 4
          ws.Cells(new_table_row, 10).Font.ColorIndex = 1
        End If
            
'check for/correct an opening value of 0, calculate and record/format percent yearly change
        If year_opening_price = 0 Then
          percent_yeary_change = 0
        Else
          percent_yearly_change = (yearly_change / year_opening_price)
        End If
        
        ws.Cells(new_table_row, 11).Value = percent_yearly_change
        ws.Cells(new_table_row, 11) = FormatPercent(ws.Cells(new_table_row, 11), 2)
                
'clear total volume for next ticker symbol, advance the row
        total_volume = 0
        new_table_row = new_table_row + 1
        
'if it's not the end or beginning of a symbol just accumulate volume
      Else
        total_volume = total_volume + ws.Cells(Row, 7).Value

      End If
    Next Row
        
'place column headers for bonus
        ws.Cells(2, 15).Value = "Greatest % Increase"
        ws.Cells(3, 15).Value = "Greatest % Decrease"
        ws.Cells(4, 15).Value = "Greatest Total Volume"
        
        ws.Cells(1, 16).Value = "Ticker"
        ws.Cells(1, 17).Value = "Value"
        
'find max of percent change, find its row and pull the ticker symbol, add both to bonus table
        lastRowK = (Cells(Rows.Count, 11).End(xlUp).Row)
        greatest_percent_increase = WorksheetFunction.Max(Range("K2:K" & lastRowK))
        ws.Cells(2, 17).Value = greatest_percent_increase
        ws.Cells(2, 17) = FormatPercent(ws.Cells(2, 17), 2)
        target_row = WorksheetFunction.Match(greatest_percent_increase, Range("K2:K" & lastRowK), 0) + 1
        ticker_symbol = ws.Cells(target_row, 9).Value
        ws.Cells(2, 16).Value = ticker_symbol
             
'find min of percent change, find its row and pull the ticker symbol, add both to bonus table
        greatest_percent_decrease = WorksheetFunction.Min(Range("K2:K" & lastRowK))
        ws.Cells(3, 17).Value = greatest_percent_decrease
        ws.Cells(3, 17) = FormatPercent(ws.Cells(3, 17), 2)
        target_row = WorksheetFunction.Match(greatest_percent_decrease, Range("K2:K" & lastRowK), 0) + 1
        ticker_symbol = ws.Cells(target_row, 9).Value
        ws.Cells(3, 16).Value = ticker_symbol
        
'find max of total volume, find its row and pull the ticker symbol, add both to bonus table
        lastRowL = (Cells(Rows.Count, 12).End(xlUp).Row)
        greatest_total_volume = WorksheetFunction.Max(Range("L2:L" & lastRowL))
        ws.Cells(4, 17).Value = greatest_total_volume
        target_row = WorksheetFunction.Match(greatest_total_volume, Range("L2:L" & lastRowK), 0) + 1
        ticker_symbol = ws.Cells(target_row, 9).Value
        ws.Cells(4, 16).Value = ticker_symbol
        
'advance the worksheet
  Next ws



End Sub




