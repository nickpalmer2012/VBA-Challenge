Attribute VB_Name = "Module1"
Sub stocks()
'Set worksheet dimension
Dim ws As Worksheet

    'Loop through each worksheet using Worksheet object
    'Set headers on each worksheet

For Each ws In Worksheets
'Ticker header
    
    ws.Cells(1, 9).Value = "Ticker"
'yearly change header
    ws.Cells(1, 10).Value = "Yearly Change"
'Percent Change header
    ws.Cells(1, 11).Value = "Percent Change"
'Total stock volume header
    ws.Cells(1, 12).Value = "Total Stock Volume"
Next ws

'Loop through all worksheets in workbook
For Each ws In Worksheets
    
    'Declare variables as appropriate data type and set inital values
    Dim ticker As String
    Dim price_change As Double
    Dim open_price As Double
    Dim close_price As Double
    Dim percent_change As Double
    Dim total_stock_volume As Double
    Dim high_ticker As String
    Dim low_ticker As String
    Dim max_percent_change As Double
    Dim min_percent_change As Double
    Dim max_volume_ticker As String
    Dim max_volume As Double
    
    'Set location of summary table on each worksheet
    Dim summary_table_row As Integer
    summary_table_row = 2
    
    'Declare row count variable
    Dim end_row As Long
    
    'Set variable value for the last row of each worksheet
     end_row = ws.Cells(Rows.Count, 1).End(xlUp).Row
     
     'Set initial value of opening price for stock on worksheet
     
     open_price = ws.Cells(2, 3).Value
     
     'Loop through each of the rows on each worksheet
     For i = 2 To end_row
     
        'set up if statement to determine if next row has different ticker name than current row
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
     
            'set up ticker name starting point
            ticker = ws.Cells(i, 1).Value
     
            'Calculate price change by subtracting the open price from the close price
            close_price = ws.Cells(i, 6).Value
            price_change = close_price - open_price
            
            'Calculate price change as a percentage
            percent_change = (price_change / open_price)
          
            
            
            'Calculate stock volume for one ticker
            total_stock_volume = total_stock_volume + ws.Cells(i, 7).Value
     
            'Set yearly change column to change color; Red if negative, green if positive
                If (price_change > 0) Then
                ws.Range("J" & summary_table_row).Interior.ColorIndex = 4
                'Else set color to red, for negative
                Else
                ws.Range("J" & summary_table_row).Interior.ColorIndex = 3
                End If
                
            'Print the ticker in column I
            ws.Range("I" & summary_table_row).Value = ticker
     
            'Print the price change in column J
            ws.Range("J" & summary_table_row).Value = price_change
     
            'Print the price change as a percentage in column K
            ws.Range("K" & summary_table_row).Value = Format(percent_change, "0.00%")
            
     
            'Print the total stock volume in column L
            ws.Range("L" & summary_table_row).Value = total_stock_volume
            
            'fetch next opening price for price change calculation
            open_price = ws.Cells(i + 1, 3).Value
            
            
     
            'Add 1 to the summary table row count
            summary_table_row = summary_table_row + 1
                
                'Complete calculations for greatest % increase, greatest % decrease
                
                'If condition for if percent change is higher than the current maximum stored value, starting with 0
                If (percent_change > max_percent_change) Then
                'set max change % variable equal to current loop's change if above conditions are met
                max_percent_change = percent_change
                'Set high ticker variable equal to current loop's ticker name if above conditons are met
                high_ticker = ticker
                
                'Else condition for if percent change is lower than the current stored minimum change value, starting with 0
                 ElseIf (percent_change < min_percent_change) Then
                 'set min % variable to current % change if above conditions are met
                 min_percent_change = percent_change
                 'Set low ticker variable to the current ticker name if above conditions are met
                 low_ticker = ticker
                 
                 End If
                 
                 
                 'Complete calculations for greatest total stock volume
                 
                 'If condition for if current stock volume is greater than the max volume that is stored
                 If (total_stock_volume > max_volume) Then
                 'Set max volume variable to current total stock volume variable if above condition is met
                 max_volume = total_stock_volume
                 'Set max ticker variable to current ticker name if above conditions are met
                 max_volume_ticker = ticker
                 
                 End If
                 
                 
     
            'Reset values for total stock volume and price change when the ticker changes
            percent_change = 0
            total_stock_volume = 0
        'Else statment to input new total stock volume
        Else
            total_stock_volume = total_stock_volume + ws.Cells(i, 7)
     
        End If
     
     Next i
     
     'Print min % change, max % change, and greatest total volume on current worksheet
     
     'Print row headers on each worksheet
     ws.Range("O2").Value = "Greatest % Increase"
     ws.Range("O3").Value = "Greatest % Decrease"
     ws.Range("O4").Value = "Greatest Total Volume"
     ws.Range("P1").Value = "Ticker"
     ws.Range("Q1").Value = "Value"
     'Print min/max values in correct cells on each worksheet
     
     'min/maxtickers
     ws.Range("P2").Value = high_ticker
     ws.Range("P3").Value = low_ticker
     ws.Range("P4").Value = max_volume_ticker
     'min max values
     ws.Range("Q2").Value = Format(max_percent_change, "0.00%")
     ws.Range("Q3").Value = Format(min_percent_change, "0.00%")
     ws.Range("Q4").Value = max_volume
     'Autofit all columns in the selected worksheet
     ws.Columns.AutoFit
     
     
     
     
Next ws
  
    
    
    
    



End Sub

