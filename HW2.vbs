Sub Calculate_Stock_Volume()

For Each ws In Worksheets

lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row

ws.Range("I1").Value = "Ticker"
ws.Range("J1").Value = "Total Stock Volume"


Dim current_ticker As String
current_ticker = ws.Cells(2, 1).Value

Dim Total_volume As Double
Total_volume = 0


Dim ticker_number As Integer
ticker_number = 1

For i = 2 To lastrow - 1
    If ws.Cells(i, 1).Value = current_ticker Then
        Total_volume = Total_volume + ws.Cells(i, 7).Value
    
    Else
        ws.Cells(ticker_number + 1, 9).Value = current_ticker
        ws.Cells(ticker_number + 1, 10).Value = Total_volume
        
        ticker_number = ticker_number + 1
    
        'reset total_volume
        Total_volume = ws.Cells(i, 7).Value
        
        'reset ticker
        current_ticker = ws.Cells(i, 1).Value
        
    End If
    
Next i

Next ws

End Sub
