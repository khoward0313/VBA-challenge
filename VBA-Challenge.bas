Attribute VB_Name = "Module1"
Sub VBA_Challenge()
    Dim ws As Worksheet
    Dim Ticker As String
    Dim Yearly_change As Double
    Dim Total_stock_volume As Double
    Dim Percent_change As Double
    Dim ticker_row As Integer
    Dim Greatest_percent_increase As Double
    Dim Greatest_percent_decrease As Double
    Dim Greatest_total_volume As Double
    Dim Open_price As Double
    Dim Last_Row As Long
    Dim Opening_price_value As Double
    Dim Summary_ticker_row As Long
    
    For Each ws In Worksheets
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Yearly Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
    
        ws.Range("O2").Value = "Greatest % Increase"
        ws.Range("O3").Value = "Greatest % Decrease"
        ws.Range("O4").Value = "Greatest Total Volume"
        ws.Range("P1").Value = "Ticker"
        ws.Range("Q1").Value = "Value"
        
        Total_stock_volume = 0
        ticker_row = 2
        Open_price = ws.Cells(2, 3).Value
        
        Last_Row = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        
        For i = 2 To Last_Row
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                Ticker = ws.Cells(i, 1).Value
                Total_stock_volume = Total_stock_volume + ws.Cells(i, 7).Value
                
                ws.Range("I" & ticker_row).Value = Ticker
                ws.Range("L" & ticker_row).Value = Total_stock_volume
                
                Yearly_change = ws.Cells(i, 6).Value - Open_price
                ws.Range("J" & ticker_row).Value = Yearly_change
                ws.Range("J" & ticker_row).NumberFormat = "$0.00"
                
                If Open_price = 0 Then
                    Percent_change = 0
                Else
                    Percent_change = Yearly_change / Open_price
                End If
                
                ws.Range("K" & ticker_row).Value = Percent_change
                ws.Range("K" & ticker_row).NumberFormat = "0.00%"
                
                Summary_ticker_row = Summary_ticker_row + 1
                Total_stock_volume = 0
                Open_price = ws.Cells(i + 1, 3).Value
                ticker_row = ticker_row + 1
            Else
                Total_stock_volume = Total_stock_volume + ws.Cells(i, 7).Value
            End If
        Next i
        
        Greatest_percent_increase = WorksheetFunction.Max(ws.Range("K2:K" & Last_Row))
        Greatest_percent_decrease = WorksheetFunction.Min(ws.Range("K2:K" & Last_Row))
        Greatest_total_volume = WorksheetFunction.Max(ws.Range("L2:L" & Last_Row))
        
        ws.Range("P2").Value = WorksheetFunction.Index(ws.Range("I2:I" & Last_Row), WorksheetFunction.Match(Greatest_percent_increase, ws.Range("K2:K" & Last_Row), 0))

    Next ws
End Sub
