Sub StockMarketHomework():

    ' loop through all worksheets
    For Each ws In Worksheets

        ' set column headers
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Yearly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Stock Volume"
        ws.Range("O2").Value = "Greatest % Increase"
        ws.Range("O3").Value = "Greatest % Decrease"
        ws.Range("O4").Value = "Greatest Total Volume"
        ws.Range("P1").Value = "Ticker"
        ws.Range("Q1").Value = "Value"

        ' declare variables
        Dim TickerName As String
        Dim LastRow As Long
        Dim TotalTickerVolume As Double
        TotalTickerVolume = 0
        
        Dim count As Long
        count = 2
        
        Dim YearlyOpen As Double
        Dim YearlyClose As Double
        Dim YearlyChange As Double
        
        Dim PreviousAmount As Long
        PreviousAmount = 2
        
        Dim PercentChange As Double
        Dim GreatestIncrease As Double
        GreatestIncrease = 0
        Dim GreatestDecrease As Double
        GreatestDecrease = 0
        
        Dim LastRowValue As Long
        
        Dim GreatestTotalVolume As Double
        GreatestTotalVolume = 0

        ' last row
        LastRow = ws.Cells(Rows.count, 1).End(xlUp).Row
        
        For i = 2 To LastRow

            ' endif for total ticket volume
            TotalTickerVolume = TotalTickerVolume + ws.Cells(i, 7).Value
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then


                ' set ticker name & add to summary table
                TickerName = ws.Cells(i, 1).Value
                ws.Range("I" & count).Value = TickerName
                ' add ticker total to summary table
                ws.Range("L" & count).Value = TotalTickerVolume
                ' reset ticker total
                TotalTickerVolume = 0

                ' declare variables for yearly open, close, & change
                YearlyOpen = ws.Range("C" & PreviousAmount)
                YearlyClose = ws.Range("F" & i)
                YearlyChange = YearlyClose - YearlyOpen
                ws.Range("J" & count).Value = YearlyChange

                ' determine percent change
                If YearlyOpen = 0 Then
                    PercentChange = 0
                Else
                    YearlyOpen = ws.Range("C" & PreviousAmount)
                    PercentChange = YearlyChange / YearlyOpen
                End If
                ' add decimal places & % sign
                ws.Range("K" & count).NumberFormat = "0.00%"
                ws.Range("K" & count).Value = PercentChange

                ' conditional formatting for yearly change
                If ws.Range("J" & count).Value >= 0 Then
                    ws.Range("J" & count).Interior.ColorIndex = 4
                Else
                    ws.Range("J" & count).Interior.ColorIndex = 3
                End If
            
                ' add 1 to summary table row
                count = count + 1
                PreviousAmount = i + 1
                End If
            Next i
            
' bonus challenge for HW

            ' greatest % increase, greatest % decrease and greatest total volume
            LastRow = ws.Cells(Rows.count, 11).End(xlUp).Row
        
            ' loop for bonus challenge information
            For i = 2 To LastRow
                If ws.Range("K" & i).Value > ws.Range("Q2").Value Then
                    ws.Range("Q2").Value = ws.Range("K" & i).Value
                    ws.Range("P2").Value = ws.Range("I" & i).Value
                End If

                If ws.Range("K" & i).Value < ws.Range("Q3").Value Then
                    ws.Range("Q3").Value = ws.Range("K" & i).Value
                    ws.Range("P3").Value = ws.Range("I" & i).Value
                End If

                If ws.Range("L" & i).Value > ws.Range("Q4").Value Then
                    ws.Range("Q4").Value = ws.Range("L" & i).Value
                    ws.Range("P4").Value = ws.Range("I" & i).Value
                End If

            Next i
        ' add two decimal places & % sign, format colums
            ws.Range("Q2").NumberFormat = "0.00%"
            ws.Range("Q3").NumberFormat = "0.00%"
            Range("Q:Q").EntireColumn.AutoFit

    Next ws

End Sub

