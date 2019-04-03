Sub stock_easy()

For Each ws In Worksheets

Dim stock_name As String
Dim close_price As String
Dim open_price As String
Dim price_change As String

Dim open_row As Long
open_row = 2

Dim ticker_row As Integer
ticker_row = 2

Dim total_volume As Double
total_volume = 0

ws.Range("I1").Value = "Stock"
ws.Range("j1").Value = "Total Volume"

ws.Columns("j:k").Insert Shift:=xlToRight

ws.Range("j1").Value = "Yearly Price Change"
ws.Range("k1").Value = "Percent Change"

ws.Range("P1").Value = "Stock"
ws.Range("Q1").Value = "Value"
ws.Range("O2").Value = "Greatest % Increase"
ws.Range("O3").Value = "Greatest % Decrease"
ws.Range("O4").Value = "Greatest Total Volume"

Max = 0

Min = 0

greatest_volume = 0


lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row

    For i = 2 To lastrow
    
    open_price = ws.Cells(open_row, 3).Value
    
    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
    
    stock_name = ws.Cells(i, 1).Value
    ws.Range("I" & ticker_row).Value = stock_name
    
    total_volume = total_volume + ws.Cells(i, 7).Value
    ws.Range("L" & ticker_row).Value = total_volume
    
        If (total_volume > greatest_volume) Then
        greatest_volume = total_volume
        ws.Range("Q4").Value = greatest_volume
        ws.Range("P4").Value = ws.Cells(i, 1).Value
        
        End If
        
    
    close_price = ws.Cells(i, 6).Value
    
    price_change = close_price - open_price
    ws.Range("j" & ticker_row).Value = price_change
    
        If (price_change > 0) Then
        ws.Range("j" & ticker_row).Interior.ColorIndex = 4
        
        Else
        ws.Range("j" & ticker_row).Interior.ColorIndex = 3
        
        End If
    
    percent_change = (open_price / close_price) - 1
    ws.Range("k" & ticker_row).Value = percent_change
    ws.Range("k" & ticker_row).Style = "Percent"
        
        If (percent_change > Max) Then
        Max = percent_change
        ws.Range("Q2").Value = Max
        ws.Range("q2").Style = "Percent"
        ws.Range("P2").Value = ws.Cells(i, 1).Value
        
        ElseIf (percent_change < Min) Then
        Min = percent_change
        ws.Range("Q3").Value = Min
        ws.Range("Q3").Style = "Percent"
        ws.Range("P3").Value = ws.Cells(i, 1).Value

        
        End If
        

    ticker_row = ticker_row + 1
    
    open_row = i + 1
    
    total_volume = 0
    
    Else
    total_volume = total_volume + ws.Cells(i, 7).Value
    
    End If
    
    Next i
    
    ws.Columns("I:Q").AutoFit
    
Next ws

End Sub


