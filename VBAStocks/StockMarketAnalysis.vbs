Sub StockMarket()

    Dim CurrentWS As Worksheet

    For Each CurrentWS In Worksheets

    CurrentWS.Activate

        Dim Ticker As String
        Dim OpenPrice, ClosePrice, YearlyChange, PercentChange As Double
        Dim RowNum, TotalStockVol, i, j, k As Long
        Dim StockRange As Range

        'Set the header for summary table
        Range("I1").Value = "Ticker"
        Range("J1").Value = "Yearly Change"
        Range("K1").Value = "Percent Change"
        Range("L1").Value = "Total Stock Volume"
        
        
        'Count number of rows
        RowNum = ActiveSheet.UsedRange.Rows.Count
        'Use j to store the first row of each stock
        j = 2
        'use k to store the row number of summary table
        k = 2

        For i = 2 To RowNum
            
            'if the next stock ticker does not equal to the current one
            If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
               
               'Return the stock ticker
                Ticker = Cells(i, 1).Value
                'Return the open price at row k (the first row of the current stock section)
                OpenPrice = Range("C" & j).Value
                'return the close price at row i (the last row of the current stock)
                ClosePrice = Range("F" & i).Value
                
                'Calculate the yearly price change and percent change
                YearlyChange = ClosePrice - OpenPrice
                
                If OpenPrice <> 0 Then
                    PercentChange = YearlyChange / OpenPrice
                Else: PercentChange = 0
                End If
                
                'set the range of the volume for current stock
                Set StockRange = Range(Cells(j, 7), Cells(i, 7))
                'Sum up the total volume of current stock
                TotalStockVol = WorksheetFunction.Sum(StockRange)
                
                'Return the results to summary table
                Range("I" & k).Value = Ticker
                Range("J" & k).Value = YearlyChange
                Range("K" & k).Value = PercentChange
                Range("L" & k).Value = TotalStockVol '
                'Format the summary table percent cells
                Range("K" & k).NumberFormat = "0.00%"
                
                If Range("K" & k).Value >= 0 Then
                    Range("K" & k).Interior.Color = vbGreen
                Else
                    Range("K" & k).Interior.Color = vbRed
                End If
                    
                'Set j to the next row, the start row of next stock
                j = i + 1
                'set k to plus one as added one row data to summary table
                k = k + 1
                
            End If
            
        Next i

        'Return greatest increase, decrease and total volume
        
        Dim GreatIn, GreatDe, InRow, DeRow, VolRow As Long
        Dim GreatStockVol As Double

        'Add headers to summary table
        Range("P1").Value = "Ticker"
        Range("Q1").Value = "Value"
        Range("O2").Value = "Greatest % Increase"
        Range("O3").Value = "Greatest % Decrease"
        Range("O4").Value = "Greatest Total Volume"
        
        'Find the greatest increase, decrease and volume
        GreatIn = WorksheetFunction.Max(Range("K2:K" & k - 1))
        GreatDe = WorksheetFunction.Min(Range("K2:K" & k - 1))
        GreatStockVol = WorksheetFunction.Max(Range("L2:L" & k - 1))
        
        'Return values to corresponding cells
        Range("Q2").Value = GreatIn
        Range("Q3").Value = GreatDe
        Range("Q4").Value = GreatStockVol
        
        'format the increase and decrease value to percentage
        Range("Q2:Q3").NumberFormat = "0.00%"
        
        'Find row numbers for the three valyes
        InRow = WorksheetFunction.Match(Range("Q2").Value, Range("K2:K" & k - 1), 0)
        DeRow = WorksheetFunction.Match(Range("Q3").Value, Range("K2:K" & k - 1), 0)
        VolRow = WorksheetFunction.Match(Range("Q4").Value, Range("L2:L" & k - 1), 0)
        
        'Return the stock tickers based on above row numbers
        Range("P2").Value = Range("I" & InRow + 1)
        Range("P3").Value = Range("I" & DeRow + 1)
        Range("P4").Value = Range("I" & VolRow + 1)
     
     Next
     
End Sub




