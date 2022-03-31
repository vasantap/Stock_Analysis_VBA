Attribute VB_Name = "Module1"
Sub StockDataAnalysis()
    'Declare variables
    Dim ws As Worksheet
    'Dim Ticker_Name As String
    'Dim Opening_Price As Double
    'Dim Closing_Price As Double
    'Dim Yearly_Change As Double
    'Dim Percent_Change As Double
    'Dim Total_Stock_Volume As Double
    
    'Loop through all the stock worksheets
    For Each ws In ActiveWorkbook.Worksheets
        ws.Activate
        LastRowWs = ws.Cells(Rows.Count, 1).End(xlUp).Row
        MsgBox (LastRowWs)
        
    Next ws
End Sub
