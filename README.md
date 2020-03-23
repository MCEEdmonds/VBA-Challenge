# VBA-Challenge

VBA Homework - The VBA of Wall Street

Background
You are well on your way to becoming a programmer and Excel master! In this homework assignment you will use VBA scripting to analyze real stock market data. Depending on your comfort level with VBA, you may choose to challenge yourself with a few of the challenge tasks.
 
# 
- Yearly change from opening price at the beginning of a given year to the closing price at the end of that year.
- The percent change from opening price at the beginning of a given year to the closing price at the end of that year.
- The total stock volume of the stock.

This is the VBA script:including one Challenge where 
"Greatest % increase", "Greatest % Decrease" and "Greatest total volume" are listed

Sub VBAHomework()

'Set a variable for holding the ticker name
        Dim tickername As String
        Dim tickervolume As Double
        tickervolume = 1

        'Keep track of the location for each ticker name in the summary table
        Dim summary_ticker_row As Integer
        summary_ticker_row = 2
        
        Dim open_price As Double
        open_price = Cells(2, 3).Value
        
        Dim close_price As Double
        Dim yearly_change As Double
        Dim percent_change As Double

        'Label the Summary Table headers
        Cells(1, 9).Value = "Ticker"
        Cells(1, 10).Value = "Yearly Change"
        Cells(1, 11).Value = "Percent Change"
        Cells(1, 12).Value = "Total Stock Volume"

        'Count the number of rows in the first column.
        lastrow = Cells(Rows.Count, 1).End(xlUp).Row


        For i = 2 To lastrow

            'Searches for when the value of the next cell is different than that of the current cell
            If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
              tickername = Cells(i, 1).Value

              'Print the ticker name, percent change, and total volume in the summary table
              tickervolume = tickervolume + Cells(i, 7).Value
              Range("I" & summary_ticker_row).Value = tickername
              Range("L" & summary_ticker_row).Value = tickervolume

              ' Assign closing price and calculate yearly change
              close_price = Cells(i, 6).Value
              yearly_change = (close_price - open_price)
              
              'Print the yearly change for each ticker in the summary table
              Range("J" & summary_ticker_row).Value = yearly_change

             'Check for the non-divisibilty condition when calculating the percent change
                If (open_price = 0) Then
                    percent_change = 0

                Else
                    percent_change = yearly_change / open_price
                End If

              'Print the yearly change for each ticker in the summary table
              Range("K" & summary_ticker_row).Value = percent_change
              Range("K" & summary_ticker_row).NumberFormat = "0.00%"
   
              'Reset the row counter. Add one to the summary_ticker_row
              summary_ticker_row = summary_ticker_row + 1

              'Reset volume of trade to zero
              tickervolume = 0

              'Reset the opening price
              open_price = Cells(i + 1, 3)
            
            Else
               'Add the volume of trade
              tickervolume = tickervolume + Cells(i, 7).Value

            End If
        
        Next i

    lastrow_summary_table = Cells(Rows.Count, 9).End(xlUp).Row
    
    'Color code yearly change
    
    For i = 2 To lastrow_summary_table
            If Cells(i, 10).Value > 0 Then
                Cells(i, 10).Interior.ColorIndex = 10
            Else
                Cells(i, 10).Interior.ColorIndex = 3
            End If
    Next i
    
    ' Assign labels to cells
    Cells(1, 15).Value = "Ticker"
    Cells(1, 16).Value = "Value"
    Cells(2, 14).Value = "Greatest % Increase"
    Cells(3, 14).Value = "Greatest % Decrease"
    Cells(4, 14).Value = "Greatest Total Volume"
    
    For i = 2 To lastrow_summary_table
            'Find the maximum percent change
            If Cells(i, 11).Value = Application.WorksheetFunction.Max(Range("K2:K" & lastrow_summary_table)) Then
                Cells(2, 15).Value = Cells(i, 9).Value
                Cells(2, 16).Value = Cells(i, 11).Value
                Cells(2, 16).NumberFormat = "0.00%"

            'Find the minimum percent change
            ElseIf Cells(i, 11).Value = Application.WorksheetFunction.Min(Range("K2:K" & lastrow_summary_table)) Then
                Cells(3, 15).Value = Cells(i, 9).Value
                Cells(3, 16).Value = Cells(i, 11).Value
                Cells(3, 16).NumberFormat = "0.00%"
            
            'Find the maximum volume of trade
            ElseIf Cells(i, 12).Value = Application.WorksheetFunction.Max(Range("L2:L" & lastrow_summary_table)) Then
                Cells(4, 15).Value = Cells(i, 9).Value
                Cells(4, 16).Value = Cells(i, 12).Value
                
            End If
            
        Next i
        
End Sub
