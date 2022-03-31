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
        Ticker_Counter = 2
        Opening_Price = 0
        Closing_Price = 0
        Yearly_Change = 0
        Percent_Change = 0
        Total_Stock_Volume = 0
        
        
        For i = 2 To LastRowWs
        'The ticker symbol
            If ws.Cells(i + 1, 1) <> ws.Cells(i, 1) Then
                ws.Cells(Ticker_Counter, 9).Value = ws.Cells(i, 1).Value
                Ticker_Counter = Ticker_Counter + 1
            End If
         Next i
        
    Next ws
End Sub
