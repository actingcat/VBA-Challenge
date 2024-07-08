# VBA-Challenge
Sub stocks()
Dim totalvolume As Double
Dim ticker As String
Dim ticker_counter, ticker_close As Double
Dim quarterly_open, quarterly_close As Double
Dim ws As Worksheet

For Each ws In Worksheets

totalvolume = 0
ticker_counter = 2
ticker_close = 2
For i = 2 To ws.Cells(Rows.Count, 1).End(xlUp).Row
    totalvolume = totalvolume + ws.Cells(i, 7).Value
    ticker = ws.Cells(i, 1).Value
    quarterly_open = ws.Cells(ticker_close, 3).Value
    
    If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then
        quarterly_close = Cells(i, 6).Value
        ws.Cells(ticker_counter, 9).Value = ticker
        ws.Cells(ticker_counter, 10).Value = quarterly_close - quarterly_open
        ws.Cells(ticker_counter, 11).Value = (quarterly_close - quarterly_open) / quarterly_open
        ws.Cells(ticker_counter, 12).Value = totalvolume
        If ws.Cells(ticker_counter, 10).Value > 0 Then
            ws.Cells(ticker_counter, 10).Interior.ColorIndex = 50
        Else
            ws.Cells(ticker_counter, 10).Interior.ColorIndex = 22
        End If
        
     totalvolume = 0
     ticker_counter = ticker_counter + 1
     ticker_close = i + 1
          
     End If
    Next i
'beginning completed with assistance from Tutor

     
Dim greatestIncrease As Double
Dim greatestDecrease As Double
Dim maxTicker As String
Dim maxTicker2 As String
Dim maxTotal As Double
maxTotal = -9999999
For i = 2 To ws.Cells(Rows.Count, 1).End(xlUp).Row
    If ws.Cells(i, 9).Value <> ws.Cells(i + 1, 9).Value Then
        If ws.Cells(i, 12).Value > maxTotal Then
            maxTotal = ws.Cells(i, 12).Value
            maxTicker = ws.Cells(i, 9).Value
        End If
    End If


If ws.Cells(i, 9).Value <> ws.Cells(i + 1, 9).Value Then
        If ws.Cells(i, 11).Value > greatestIncrease Then
            greatestIncrease = ws.Cells(i, 11).Value
            maxTicker2 = ws.Cells(i, 9).Value
        End If
    End If
    
If ws.Cells(i, 9).Value <> ws.Cells(i + 1, 9).Value Then
        If ws.Cells(i, 11).Value < greatestDecrease Then
            greatestDecrease = ws.Cells(i, 11).Value
            maxTicker3 = ws.Cells(i, 9).Value
            End If
        End If
    
Next i

ws.Range("Q2").Value = greatestIncrease
ws.Range("P4").Value = maxTicker
ws.Range("P2").Value = maxTicker2
ws.Range("Q4").Value = maxTotal
ws.Range("P3").Value = maxTicker3
ws.Range("Q3").Value = greatestDecrease
ws.Range("P2", "P4").Interior.ColorIndex = 38
ws.Range("Q2", "Q4").Interior.ColorIndex = 39
ws.Range("O2", "O4").Interior.ColorIndex = 24

Next ws

End Sub
