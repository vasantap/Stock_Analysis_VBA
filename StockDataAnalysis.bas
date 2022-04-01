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
        
        'Initializing variables
        
        'Counter to find the open price and its ticker
        Ticker_Counter_Start = 2
        Opening_Price = 0
        Yearly_Change = 0
        Percent_Change = 0
        'Counter to find the close price of the corresponding open price
        Ticker_Counter_End = 2
        Total_Stock_Volume = 0
        
        For i = 2 To LastRowWs
                        
            'The ticker symbol
            If ws.Cells(i, 1) <> ws.Cells(i + 1, 1) Then
                ws.Cells(Ticker_Counter_Start, 9).Value = ws.Cells(i, 1).Value
               
                'Yearly Change
                ws.Cells(Ticker_Counter_Start, 10).Value = ws.Cells(i, 6).Value - ws.Cells(Ticker_Counter_End, 3).Value
                Yearly_Change = ws.Cells(Ticker_Counter_Start, 10).Value
                
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
                 Total_Stock_Volume = Total_Stock_Volume + Cells(i, 7).Value
                Cells(Ticker_Counter_End, 12).Value = Total_Stock_Volume
                Ticker_Counter_Start = Ticker_Counter_Start + 1
                Total_Stock_Volume = 0
                Ticker_Counter_End = i + 1
              Else
                Total_Stock_Volume = Total_Stock_Volume + Cells(i, 7).Value
             End If
        Next i
    Next ws
End Sub
