Attribute VB_Name = "Module1"
Sub stock_loop()
'Loop through all worksheets

    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
    
    
'Set a variable for holding the ticker name, the column of interest
        Dim tickername As String
    
'Set a variable for holding a total count on the total volume of trade
        Dim tickervolume As Double
        tickervolume = 0

'Keep track of the location for each ticker name in the summary table
        Dim summary_ticker_row As Integer
        summary_ticker_row = 2
        
'Set initial open_price
        Dim open_price As Double
        open_price = ws.Cells(2, 3).Value

'Declare variables
        Dim close_price As Double
        Dim quarterly_change As Double
        Dim percent_change As Double
 
'Label the Summary Table headers
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Quarterly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"

'Count the number of rows in the first column.
        lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row


        'Loop through the rows by the ticker names
        
        For i = 2 To lastrow

            'Searches for when the value of the next cell is different than that of the current cell
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
        
              'Set the ticker name
              tickername = ws.Cells(i, 1).Value

              'Add the volume of trade
              tickervolume = tickervolume + ws.Cells(i, 7).Value

              'Print the ticker name in the summary table
              ws.Range("I" & summary_ticker_row).Value = tickername

              'Print the trade volume for each ticker in the summary table
              ws.Range("L" & summary_ticker_row).Value = tickervolume

              'Set the closing price
              close_price = ws.Cells(i, 6).Value

              'Calculate quarterly change
               quarterly_change = (close_price - open_price)
              
              'Print the quarterly change for each ticker in the summary table
              ws.Range("J" & summary_ticker_row).Value = quarterly_change

              'Check for the non-divisibilty condition when calculating the percent change
                If open_price = 0 Then
                    percent_change = 0
                Else
                    percent_change = quarterly_change / open_price
                End If

              'Print the percent change for each ticker in the summary table
              ws.Range("K" & summary_ticker_row).Value = percent_change

          'Format percent change column
              ws.Range("K" & summary_ticker_row).NumberFormat = "0.00%"
   
              'Add one to the summary_ticker_row
              summary_ticker_row = summary_ticker_row + 1

              'Reset volume of trade to zero
              tickervolume = 0

              'Reset the opening price
              open_price = ws.Cells(i + 1, 3)
            
            Else
              
               'Add the volume of trade
              tickervolume = tickervolume + ws.Cells(i, 7).Value

            
            End If
        
        Next i

    
    'Find the last row of the summary table

    lastrow_summary_table = ws.Cells(Rows.Count, 9).End(xlUp).Row
    
    

    'Loop from second row to lastrow_summary_table

        For i = 2 To lastrow_summary_table

    'if quarterly row is positive, cell interior will be green

            If ws.Cells(i, 10).Value > 0 Then

                ws.Cells(i, 10).Interior.ColorIndex = 4

            Else

    'if quarterly row is negative, cell interior will be red

                ws.Cells(i, 10).Interior.ColorIndex = 3

            End If

        Next i

    
    'Label cells

        ws.Cells(2, 15).Value = "Greatest % Increase"
        ws.Cells(3, 15).Value = "Greatest % Decrease"
        ws.Cells(4, 15).Value = "Greatest Total Volume"
        ws.Cells(1, 16).Value = "Ticker"
        ws.Cells(1, 17).Value = "Value"


    'Declare variables
        Dim Greatest_Increase As Double
        Dim Greatest_Decrease As Double
        Dim Tickerindex As Integer
        Dim PercentChangeRange As Range
        Dim GreatestTotalVolRange As Range

        

        ' Set the range where the max and min percent change will be found
        Set PercentChangeRange = ws.Range("K2", "K" & lastrow_summary_table)

        ' Find the greatest increase using the Max function
        Greatest_Increase = WorksheetFunction.Max(PercentChangeRange)

        ' Use the Match function to find the cell that hold the greatest increase
        Tickerindex = WorksheetFunction.Match(Greatest_Increase, PercentChangeRange, 0)

        '  Print the greatest increase & corresponding ticker symbol in the summary table
        ws.Range("Q2").Value = Greatest_Increase
 
        ws.Range("P2").Value = ws.Range("I" & Tickerindex + 1).Value
 
        ' Find the greatest increase using the Min function
        Greatest_Decrease = WorksheetFunction.Min(PercentChangeRange)


        ' Use the Match function to find the cell that hold the greatest Decrease
        Tickerindex = WorksheetFunction.Match(Greatest_Decrease, PercentChangeRange, 0)

        ' Print the greatest Decrease & corresponding ticker symbol in the summary tabl
        ws.Range("Q3").Value = Greatest_Decrease
 
        ws.Range("P3").Value = ws.Range("I" & Tickerindex + 1).Value
 
 
 
        ' Set the range where the greatest total volume will be found
        Set GreatestTotalVolRange = ws.Range("L2", "L" & lastrow_summary_table)
 
        ' Find the greatest total volume increase using the Max function
        Greatest_Total_Volume = WorksheetFunction.Max(GreatestTotalVolRange)
 
        ' Use the Match function to find the cell that hold the greatest total volume
        Tickerindex = WorksheetFunction.Match(Greatest_Total_Volume, GreatestTotalVolRange, 0)
 
 
        ' Print the greatest Decrease & corresponding ticker symbol in the summary tabl
            ws.Range("Q4").Value = Greatest_Total_Volume
 
            ws.Range("P4").Value = ws.Range("I" & Tickerindex + 1).Value


 
         'Format the output summary of Greatest_Increase &  to percentage
            ws.Range("Q2:Q3").NumberFormat = "0.00%"

        
    
    
            Next ws
        

End Sub
