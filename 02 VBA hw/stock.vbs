Sub Stock()
Dim ticker As String
Dim ticker_counter As Double
Dim yearly_open As Double
Dim yearly_close As Double
Dim yearly_counter as Double
Dim total_volume As Double
    
    ticker_counter = 2 
    yearly_counter = 2 
    total_volume = 0
    
    Cells(1,9).Value = "Ticker"
    Cells(1,10).Value = "Yearly Change"
    Cells(1,11).Value = "Percent Change"
    Cells(1,12).Value = "Total Stock Volume"
    
    For i = 2 To Cells(Rows.Count, 1).End(xlUp).Row

        total_volume = total_volume + Cells(i, 7).Value
        ticker = Cells(i, 1).Value
        yearly_open = Cells(yearly_counter, 3)
        
        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
            yearly_end = Cells(i, 6)
            Cells(ticker_counter, 9).Value = ticker
            Cells(ticker_counter, 10).Value = yearly_end - yearly_open
            If yearly_open = 0 Then
                Cells(ticker_counter, 11).Value = Null
            Else
                Cells(ticker_counter, 11).Value = (yearly_end - yearly_open) / yearly_open
            End If
            Cells(ticker_counter, 12).Value = total_volume
            If Cells(ticker_counter, 10).Value > 0 Then
                Cells(ticker_counter, 10).Interior.ColorIndex = 4
            Else
                Cells(ticker_counter, 10).Interior.ColorIndex = 3
            End If
            
            Cells(ticker_counter, 11).NumberFormat = "0.00%"
            total_volume = 0
            ticker_counter = ticker_counter + 1
            yearly_counter = i + 1
        End If
        
    Next i

    Columns("J").Autofit
    Columns("K").Autofit
    Columns("L").Autofit

End Sub

