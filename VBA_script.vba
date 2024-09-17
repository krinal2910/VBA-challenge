Sub QuarterlyStockAnalysis()
    Dim ws As Worksheet
    Dim lastRow As Long, i As Long
    Dim ticker As String
    Dim openPrice As Double, closePrice As Double
    Dim quarterlyChange As Double, percentChange As Double
    Dim totalVolume As Double
    Dim summaryRow As Long
    
    ' Loop through each quarter sheet
    For Each ws In ThisWorkbook.Worksheets
        If ws.Name Like "Q*" Then
            lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
            summaryRow = 2 ' Start summary output from row 2
            
            ' Clear previous summary data
            ws.Range("H2:K" & ws.Rows.Count).ClearContents
            
            ' Add headers for summary
            ws.Cells(1, 8).Value = "Ticker"
            ws.Cells(1, 9).Value = "Quarterly Change"
            ws.Cells(1, 10).Value = "Percent Change"
            ws.Cells(1, 11).Value = "Total Stock Volume"
            
            ' Variables to track greatest values
            Dim maxPercentIncrease As Double
            Dim maxPercentDecrease As Double
            Dim maxTotalVolume As Double
            Dim maxIncreaseTicker As String
            Dim maxDecreaseTicker As String
            Dim maxVolumeTicker As String
            
            ' Initialize tracking variables
            maxPercentIncrease = -1
            maxPercentDecrease = 1
            maxTotalVolume = 0
            
            i = 2
            Do While i <= lastRow
                ticker = ws.Cells(i, 1).Value
                openPrice = ws.Cells(i, 3).Value
                totalVolume = 0
                
                ' Find the last row for the current ticker
                Do While i <= lastRow And ws.Cells(i, 1).Value = ticker
                    totalVolume = totalVolume + ws.Cells(i, 7).Value
                    i = i + 1
                Loop
                
                ' Get the closing price of the last entry for the ticker
                closePrice = ws.Cells(i - 1, 6).Value
                
                ' Calculate the changes
                quarterlyChange = closePrice - openPrice
                If openPrice <> 0 Then
                    percentChange = (quarterlyChange / openPrice) * 100
                Else
                    percentChange = 0
                End If
                
                ' Output the summary data
                ws.Cells(summaryRow, 8).Value = ticker
                ws.Cells(summaryRow, 9).Value = Format(quarterlyChange, "0.00")
                ws.Cells(summaryRow, 10).Value = Format(percentChange, "0.00") & "%"
                ws.Cells(summaryRow, 11).Value = totalVolume
                
                ' Apply color formatting
                If quarterlyChange >= 0 Then
                    ws.Cells(summaryRow, 9).Interior.Color = RGB(0, 255, 0) ' Green for positive change
                Else
                    ws.Cells(summaryRow, 9).Interior.Color = RGB(255, 0, 0) ' Red for negative change
                End If
            
                
                ' Track greatest values
                If percentChange > maxPercentIncrease Then
                    maxPercentIncrease = percentChange
                    maxIncreaseTicker = ticker
                End If
                
                If percentChange < maxPercentDecrease Then
                    maxPercentDecrease = percentChange
                    maxDecreaseTicker = ticker
                End If
                
                If totalVolume > maxTotalVolume Then
                    maxTotalVolume = totalVolume
                    maxVolumeTicker = ticker
                End If
                
                ' Increment summary row
                summaryRow = summaryRow + 1
            Loop
            
            ' Output the greatest values
            ws.Cells(2, 15).Value = "Greatest % Increase"
            ws.Cells(2, 16).Value = maxIncreaseTicker
            ws.Cells(2, 17).Value = Format(maxPercentIncrease, "0.00") & "%"
            
            ws.Cells(3, 15).Value = "Greatest % Decrease"
            ws.Cells(3, 16).Value = maxDecreaseTicker
            ws.Cells(3, 17).Value = Format(maxPercentDecrease, "0.00") & "%"
            
            ws.Cells(4, 15).Value = "Greatest Total Volume"
            ws.Cells(4, 16).Value = maxVolumeTicker
            ws.Cells(4, 17).Value = maxTotalVolume
        End If
    Next ws
End Sub


