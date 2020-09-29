Sub stock()

' The ticker symbol.
' Yearly change from opening price at the beginning of a given year to the closing price at the end of that year.
' The percent change from opening price at the beginning of a given year to the closing price at the end of that year.
' The total stock volume of the stock.


    Dim r, fRow, ticker As Long
    Dim opening, closing, yearlyChange, total_stock_volume, maxIncrease, maxDecrease, maxVolume As Double
    
    
    ' Loop through all sheets
    For ws = 1 To Worksheets.Count
        
        Worksheets(ws).Select
    
        Range("I1").Value = "Ticker"
        Range("J1").Value = "Yearly Change"
        Range("K1").Value = "Percent Change"
        Range("L1").Value = "Total Stock Volume"
        
        lastRow = Cells(Rows.Count, 1).End(xlUp).Row
    
        fRow = 2
        ticker = 1
    
        For r = 2 To lastRow
    
            If Cells(r, 1).Value <> Cells(r + 1, 1).Value Then
    
                ticker = ticker + 1

                ' Ticker
                Range("I" + CStr(ticker)).Value = Cells(r, 1).Value
    
                ' Opening, Closing, and Yearly Change
                opening = Cells(fRow, 3).Value
                closing = Cells(r, 6).Value
                yearlyChange = closing - opening
                Range("J" + CStr(ticker)).Value = yearlyChange
    
                If yearlyChange <= 0 Then
                    'negative change in red = 3
                    Range("J" + CStr(ticker)).Interior.ColorIndex = 3
                Else
                    'positive change in green = 4
                    Range("J" + CStr(ticker)).Interior.ColorIndex = 4
                End If
    
                ' Conditional for indeterminate division
                If opening = 0 Then
                    Range("K" + CStr(ticker)).Value = 0
                Else
                    Range("K" + CStr(ticker)).Value = (yearlyChange / opening)
                End If
    
                ' Formated cell for percentage values
                Range("K" + CStr(ticker)).NumberFormat = "0.00%"
    
                ' Total Stock Volume
                Range("L" + CStr(ticker)).Value = total_stock_volume + Cells(r, 7).Value
    
                fRow = r + 1
                total_stock_volume = 0
    
            Else
                total_stock_volume = total_stock_volume + Cells(r, 7).Value
    
            End If
        Next r
    
        Range("P1").Value = "Ticker"
        Range("Q1").Value = "Value"
        Range("O2").Value = "Greatest % Increase"
        Range("O3").Value = "Greatest % Decrease"
        Range("O4").Value = "TGreatest Total Volume"
    
        lastRow = Cells(Rows.Count, 9).End(xlUp).Row
        
        ' Max Increase and Decrease; Max Volume per year
        Range("Q2").Formula = "=MAX(K2:K" & lastRow & ")"
        Range("Q3").Formula = "=MIN(K2:K" & lastRow & ")"
        Range("Q4").Formula = "=MAX(L2:L" & lastRow & ")"
        
        maxIncrease = Range("Q2").Value
        maxDecrease = Range("Q3").Value
        maxVolume = Range("Q4").Value
        
        ' Looking up for the values in the columns
        For r = 2 To lastRow
        
            If maxIncrease = Cells(r, 11).Value Then
                Range("P2").Value = Cells(r, 9).Value
            ElseIf maxDecrease = Cells(r, 11).Value Then
                Range("P3").Value = Cells(r, 9).Value
            End If
            
            If maxVolume = Cells(r, 12).Value Then
                Range("P4").Value = Cells(r, 9).Value
            End If
            
        Next r
        
        Columns("I:Q").EntireColumn.AutoFit

    Next ws
    

End Sub

