' Create new columns called Ticker, Quarterly Change, Percent Change, Total Stock Volume.

'Add Column Ticker
'Add Column Quarterly Change
'Add Column Percent Change
'Add Column Total Stock Volume

'Find unique tickers and add them to Ticker column
    'NextCells loop to find unique ticker name

'Calculate quarterly change for the unique ticker
    'Find First Open Cell price and last Close Price for Ticker
    'Calculate (ClosePrice-OpenPrice)
    'Name QuarterChange
    'Add to Column H
    'Add conditional formatting - Green for increase, red for decrease - Use Interior.IndexColor
    
'Calculate percent change for the unique ticker
    'Find First Open Cell price and last Close Price for Ticker
    'Calculate (QuarterChange)/OpenPrice*100
    'Add to Column I
    'Add percent format
    
'Calculate the total stock volume for the unique ticker
    'Add Column G for each ticker
    
'Run script across all 4 Worksheets

Sub StockAnalysis()

Dim ws As Worksheet

'Loop through all sheets
For Each ws In Worksheets

'Variable to hold the ticker symbol
Dim ticker As String

'Variables for finding first open and last close price
Dim firstopen As Double
Dim lastclose As Double

'Variable to hold the final values
Dim quarterChange As Double
Dim percentChange As Double
Dim totalVolume As Double

'Variables for tracking greatest values
Dim maxPercentIncrease As Double
Dim maxPercentIncreaseTicker As String
Dim maxPercentDecrease As Double
Dim maxPercentDecreaseTicker As String
Dim maxTotalVolume As Double
Dim maxTotalVolumeTicker As String

'Track location of ticker symbol in table
Dim tickerRow As Long
tickerRow = 2 'start on row 2
totalVolume = 0
firstopen = 0
maxPercentIncrease = -1
maxPercentDecrease = 1
maxTotalVolume = 0

'Add Column Names
ws.Range("H1").Value = "Ticker"
ws.Range("I1").Value = "Quarterly Change"
ws.Range("J1").Value = "Percent Change"
ws.Range("K1").Value = "Total Stock Volume"
ws.Range("O1").Value = "Ticker"
ws.Range("P1").Value = "Value"
ws.Range("N2").Value = "Greatest % Increase"
ws.Range("N3").Value = "Greatest % Decrease"
ws.Range("N4").Value = "Greatest Total Volume"

'Set up last row
lastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

    'Loop through ticker symbols
    For i = 2 To lastRow
    
        If ws.Cells(i - 1, 1).Value <> ws.Cells(i, 1).Value Then
            firstopen = ws.Cells(i, 3).Value
        End If
    
        'Calculate total volume
        totalVolume = totalVolume + ws.Cells(i, 7).Value
    
        'Check the ticker symbol is different
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            'Set the ticker name
            ticker = ws.Cells(i, 1).Value
        
            'Find last close price (from the current row, which is the last occurrence)
            lastclose = ws.Cells(i, 6).Value
        
            'Calculate QuarterlyChange
            quarterChange = lastclose - firstopen
            
            'Calculate percentChange
            percentChange = quarterChange / firstopen
            
            'Calculate percentChange
            If firstopen <> 0 Then
                percentChange = (quarterChange / firstopen)
            Else
                percentChange = 0
            End If
        
            'Print ticker to Column H
            ws.Range("H" & tickerRow).Value = ticker
            ws.Range("I" & tickerRow).Value = quarterChange
            ws.Range("J" & tickerRow).Value = percentChange
            ws.Range("J" & tickerRow).NumberFormat = "0.00%"
            ws.Range("K" & tickerRow).Value = totalVolume
        
            'Apply conditional formatting
            If quarterChange < 0 Then
                ws.Range("I" & tickerRow).Interior.ColorIndex = 3
            ElseIf quarterChange > 0 Then
                ws.Range("I" & tickerRow).Interior.ColorIndex = 4
            Else
                ws.Range("I" & tickerRow).Interior.ColorIndex = None
            End If
            
            'Track the greatest percentage increase
            If percentChange > maxPercentIncrease Then
                maxPercentIncrease = percentChange
                maxPercentIncreaseTicker = ticker
            End If

            'Track the greatest percentage decrease
            If percentChange < maxPercentDecrease Then
                maxPercentDecrease = percentChange
                maxPercentDecreaseTicker = ticker
            End If

            'Track the greatest total volume
            If totalVolume > maxTotalVolume Then
                maxTotalVolume = totalVolume
                maxTotalVolumeTicker = ticker
            End If

            'Add 1 to the tickerRow
            tickerRow = tickerRow + 1

            'Reset totalVolume for next ticker
            totalVolume = 0
            firstopen = 0
            
        End If

    Next i
    
    ' Write the results for greatest values in columns N to P
    ws.Range("O2").Value = maxPercentIncreaseTicker
    ws.Range("P2").Value = maxPercentIncrease
    ws.Range("P2").NumberFormat = "0.00%"

    ws.Range("O3").Value = maxPercentDecreaseTicker
    ws.Range("P3").Value = maxPercentDecrease
    ws.Range("P3").NumberFormat = "0.00%"

    ws.Range("O4").Value = maxTotalVolumeTicker
    ws.Range("P4").Value = maxTotalVolume
    
    'Autofit all the columns
    ws.Columns("A:P").AutoFit
    
Next ws
    
    'Message Box saying the code has finished running
    MsgBox ("Analysis Complete")
    
End Sub
