Attribute VB_Name = "Module2"
Sub StockDataAnalysis()

    'Declare variables
    Dim ws As Worksheet
    Dim Ticker_Counter_Start As Double
    Dim Ticker_Counter_End As Double
    Dim Opening_Price As Double
    Dim Yearly_Change As Double
    Dim Percent_Change As Double
    Dim Total_Stock_Volume As Double
    
    'Loop through all the stock worksheets
    For Each ws In ActiveWorkbook.Worksheets
        ws.Activate
        LastRowWs = ws.Cells(Rows.Count, 1).End(xlUp).Row
                      
        'Creating column headers
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Yearly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Stock Volume"
        
        'BONUS
        'Add script to return the stock with Greatest % increase, Greatest % decrease and Greatest total volume
        ws.Range("O2").Value = "Greatest % Increase"
        ws.Range("O3").Value = "Greatest % Decrease"
        ws.Range("O4").Value = "Greatest Total Volume"
        ws.Range("P1").Value = "Ticker"
        ws.Range("Q1").Value = "Value"
        
        'Initializing variables
        
        'Counter to find the opening price and its ticker
        Ticker_Counter_Start = 2
        Opening_Price = 0
        Yearly_Change = 0
        Percent_Change = 0
        'Counter to find the closing price of the corresponding open price
        Ticker_Counter_End = 2
        Total_Stock_Volume = 0
        
        'BONUS
        'Initialize variables and set values of variables initially to the first row in the list.
        greatest_percent_increase = ws.Cells(2, 11).Value
        greatest_percent_increase_ticker = ws.Cells(2, 9).Value
        greatest_percent_decrease = ws.Cells(2, 11).Value
        greatest_percent_decrease_ticker = ws.Cells(2, 9).Value
        greatest_stock_volume = ws.Cells(2, 12).Value
        greatest_stock_volume_ticker = ws.Cells(2, 9).Value
        
        For i = 2 To LastRowWs
                        
            'The ticker symbol
            If ws.Cells(i + 1, 1) <> ws.Cells(i, 1) Then
                ws.Cells(Ticker_Counter_Start, 9).Value = ws.Cells(i, 1).Value
                
                'Yearly Change
                ws.Cells(Ticker_Counter_Start, 10).Value = ws.Cells(i, 6).Value - ws.Cells(Ticker_Counter_End, 3).Value
                Yearly_Change = ws.Cells(Ticker_Counter_Start, 10).Value
                ws.Cells(Ticker_Counter_Start, 10).Value = Format(Yearly_Change, "###0.00")
                
                'Conditional formatting showing -ve values highlighted in red and +ve in green
                If ws.Cells(Ticker_Counter_Start, 10).Value < 0 Then
                    ws.Cells(Ticker_Counter_Start, 10).Interior.ColorIndex = 3
                Else
                    ws.Cells(Ticker_Counter_Start, 10).Interior.ColorIndex = 4
                End If
  
                'Percentage change after checking for denominator not zero
                Opening_Price = ws.Cells(Ticker_Counter_End, 3).Value
                If ws.Cells(Ticker_Counter_End, 3).Value <> 0 Then
                    Percent_Change = Yearly_Change / Opening_Price
                    'Format with %
                    ws.Cells(Ticker_Counter_Start, 11).Value = Format(Percent_Change, "Percent")
                Else
                    ws.Cells(Ticker_Counter_Start, 11).Value = Format(0, "Percent")
                End If
                
                'Total stock volume is sum of block of cells from ticker counter start to ticker counter end
                Total_Stock_Volume = Total_Stock_Volume + ws.Cells(i, 7).Value
                ws.Cells(Ticker_Counter_Start, 12).Value = Total_Stock_Volume
                Ticker_Counter_Start = Ticker_Counter_Start + 1
                Ticker_Counter_End = i + 1
                Total_Stock_Volume = 0
             Else
                Total_Stock_Volume = Total_Stock_Volume + ws.Cells(i, 7).Value
                ws.Cells(Ticker_Counter_Start, 12).Value = Total_Stock_Volume
             End If
             
            'Iterate through columns I and K to calculate
            'Ticker with greatest % increase
            If ws.Cells(i, 11).Value > greatest_percent_increase Then
                greatest_percent_increase = ws.Cells(i, 11).Value
                greatest_percent_increase_ticker = ws.Cells(i, 9).Value
            End If
            
            'Ticker with greatest % decrease
            If ws.Cells(i, 11).Value < greatest_percent_decrease Then
                greatest_percent_decrease = ws.Cells(i, 11).Value
                greatest_percent_decrease_ticker = ws.Cells(i, 9).Value
            End If
        
            'Ticker with the greatest stock volume.
            If ws.Cells(i, 12).Value > greatest_stock_volume Then
                greatest_stock_volume = ws.Cells(i, 12).Value
                greatest_stock_volume_ticker = ws.Cells(i, 9).Value
            End If
        Next i
        
        'BONUS
        'Show on worksheet.
        Range("P2").Value = Format(greatest_percent_increase_ticker, "Percent")
        Range("Q2").Value = Format(greatest_percent_increase, "Percent")
        Range("P3").Value = Format(greatest_percent_decrease_ticker, "Percent")
        Range("Q3").Value = Format(greatest_percent_decrease, "Percent")
        Range("P4").Value = greatest_stock_volume_ticker
        Range("Q4").Value = greatest_stock_volume
    Next ws
End Sub

