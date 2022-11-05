Sub stocksummary():

    Dim w As Integer
    Dim i As Double
    Dim s As Integer
    Dim h As Double
     
        
    For w = 1 To 3
    
        Dim Acc_Volume As Double
        Dim yearly_change As Double
        Dim opening_price As Double
        Dim closing_price As Double
        Dim price_difference As Double
        Dim percent_change As Double
        s = 2
    
        Worksheets(w).Range("J1").Value = "Ticker Symbol"
        Worksheets(w).Range("K1").Value = "Yearly Change from the Opening Price to the Closing Price"
        Worksheets(w).Range("L1").Value = "Percent Change from the Opening Price to the Closing Price"
        Worksheets(w).Range("M1").Value = "Total Stock Volume"
        
        last_row = Worksheets(w).Cells(Rows.Count, 1).End(xlUp).Row
        Acc_Volume = Worksheets(w).Cells(2, 7).Value
        opening_price = Worksheets(w).Cells(2, 3).Value
        closing_price = 0
        price_difference = 0
        
        Worksheets(w).Range("O2").Value = "Greatest % Increase"
        Worksheets(w).Range("O3").Value = "Greatest % Decrease"
        Worksheets(w).Range("O4").Value = "Greatest Total Volume"
        Worksheets(w).Range("P1").Value = "Ticker"
        Worksheets(w).Range("Q1").Value = "Value"
        ' Bonus summary table
        
        summary_lastrow = Worksheets(w).Cells(Rows.Count, 1).End(xlUp).Row
        Dim maxpercent As Double
        Dim minpercent As Double
        Dim maxtotal As Double
        Dim tickervalue As String
        Dim tickervalue2 As String
        Dim tickervalue3 As String
    
     
        maxpercent = Worksheets(w).Cells(2, 12).Value
        minpercent = Worksheets(w).Cells(2, 12).Value
        maxtotal = Worksheets(w).Cells(2, 13).Value
        tickervalue = Worksheets(w).Cells(2, 10).Value
        tickervalue2 = Worksheets(w).Cells(2, 10).Value
        tickervalue3 = Worksheets(w).Cells(2, 10).Value
    
        For i = 2 To last_row
        
                        
            If Worksheets(w).Cells(i, 1) <> Worksheets(w).Cells(i + 1, 1) Then
                
                Worksheets(w).Cells(s, 10).Value = Worksheets(w).Cells(i, 1).Value
                ' Records the ticker
                
                Worksheets(w).Cells(s, 13).Value = Acc_Volume
                ' Records the accumulated volume
                
                closing_price = Worksheets(w).Cells(i, 6).Value
                price_difference = closing_price - opening_price
                Worksheets(w).Cells(s, 11).Value = price_difference
                If Worksheets(w).Cells(s, 11).Value > 0 Then
                    Worksheets(w).Cells(s, 11).Interior.Color = vbGreen
                ElseIf Worksheets(w).Cells(s, 11).Value < 0 Then
                    Worksheets(w).Cells(s, 11).Interior.Color = vbRed
                End If
                ' Records the price difference
                
                
                percent_change = (price_difference / opening_price)
                Worksheets(w).Cells(s, 12).Value = percent_change
                Worksheets(w).Range("L" & s).Style = "Percent"
                'Records the percentage change
                                               
                
                s = s + 1
                ' Summary table count
                Acc_Volume = Worksheets(w).Cells(i + 1, 7).Value
                ' Accumulated volume reset
                opening_price = Worksheets(w).Cells(i + 1, 3).Value
                ' Declare a new opening price
                
            Else
                Acc_Volume = Acc_Volume + Worksheets(w).Cells(i + 1, 7).Value
                ' Accumulates volume (i + i+1)
                
            End If
        Next i
        
        
            For h = 3 To summary_lastrow
                If maxpercent < Worksheets(w).Cells(h, 12).Value Then
                     maxpercent = Worksheets(w).Cells(h, 12).Value
                     tickervalue = Worksheets(w).Cells(h, 10).Value
                End If
            Next h
        
            For h = 3 To summary_lastrow
                If minpercent > Worksheets(w).Cells(h, 12).Value Then
                    minpercent = Worksheets(w).Cells(h, 12).Value
                    tickervalue2 = Worksheets(w).Cells(h, 10).Value
                End If
            Next h
        
            For h = 3 To summary_lastrow
                If maxtotal < Worksheets(w).Cells(h, 13).Value Then
                    maxtotal = Worksheets(w).Cells(h, 13).Value
                    tickervalue3 = Worksheets(w).Cells(h, 10).Value
                End If
            Next h
        
        
            Worksheets(w).Range("P2").Value = tickervalue
            Worksheets(w).Range("Q2").Value = maxpercent
            Worksheets(w).Range("Q2").Style = "Percent"
            Worksheets(w).Range("P3").Value = tickervalue2
            Worksheets(w).Range("Q3").Value = minpercent
            Worksheets(w).Range("Q3").Style = "Percent"
            Worksheets(w).Range("P4").Value = tickervalue3
            Worksheets(w).Range("Q4").Value = maxtotal
        
    Next w
    



 
End Sub














