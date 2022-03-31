Attribute VB_Name = "Module1"
Sub StockDataAnalysis()
    'Declare variables
    Dim ws As Worksheet
    Dim Ticker_Name As String
    Dim Ticker_Counter As Double
    Dim Opening_Price As Double
    Dim Closing_Price As Double
    Dim Yearly_Change As Double
    Dim Percent_Change As Double
    Dim Total_Stock_Volume As Double
    
    'Loop through all the stock worksheets
    For Each ws In ActiveWorkbook.Worksheets
        ws.Activate
        LastRowWs = ws.Cells(Rows.Count, 1).End(xlUp).Row
        'MsgBox (LastRowWs)
        
        'Creating column headers
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Yearly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Stock Volume"
        
        'Initializing variables
        Ticker_Name = " "
        Ticker_Counter_Start = 2 'Counter to find the open price and its ticker
        Yearly_Change = 0
        Percent_Change = 0
        Total_Stock_Volume = 0
        Ticker_Counter_End = 2 'Counter to find the close price of the corresponding open price
        
        For i = 2 To LastRowWs
            'The ticker symbol
            If ws.Cells(i, 1) <> ws.Cells(i + 1, 1) Then
                ws.Cells(Ticker_Counter_Start, 9).Value = ws.Cells(i, 1).Value
                'Yearly Change
                ws.Cells(Ticker_Counter_Start, 10).Value = ws.Cells(i, 6).Value - ws.Cells(Ticker_Counter_End, 3).Value
        
        
             Ticker_Counter_Start = Ticker_Counter_Start + 1
             Ticker_Counter_End = i + 1
            End If
         Next i
        
    Next ws
End Sub
