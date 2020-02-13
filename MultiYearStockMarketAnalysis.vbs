Sub MultiYearStockMarketAnalysis():

For Each ws In Worksheets

lastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
lastcolumn = ws.Cells(1, Columns.Count).End(xlToLeft).Column

'define the starting row where summary information needs to be displayed

startingSummaryRow = 2

'define a counter for total stock volume per ticker
TotalStockVolume = 0

'define the headers for Summary row
ws.Cells(1, lastcolumn + 2).Value = "Ticker"
ws.Cells(1, lastcolumn + 3).Value = "Yearly Change"
ws.Cells(1, lastcolumn + 4).Value = "Percent Change"
ws.Cells(1, lastcolumn + 5).Value = "Total Stock Volume"

'counter for Open Price
OpenPriceIndex = 2

'loop through the entire list
For i = 2 To lastRow

'check if the next row has same Credit Card type, if not
If ws.Cells(i, 1) <> ws.Cells(i + 1, 1) Then

'find the Open price per ticker at the beginning of the year
YearBegOpenPrice = ws.Cells(OpenPriceIndex, 3)


'increment openpriceindex
OpenPriceIndex = i + 1

'display in the summary table
ws.Cells(startingSummaryRow, 9) = ws.Cells(i, 1).Value

'calculate the total stock volume
TotalStockVolume = TotalStockVolume + ws.Cells(i, 7).Value
ws.Cells(startingSummaryRow, 12) = TotalStockVolume

'find the Close price per ticker at the end of the year
YearEndClosePrice = ws.Cells(i, 6).Value


'Find yearly change
ws.Cells(startingSummaryRow, 10) = YearEndClosePrice - YearBegOpenPrice

'conditional formatting for color coding

If ws.Cells(startingSummaryRow, 10) >= 0 Then

ws.Cells(startingSummaryRow, 10).Interior.ColorIndex = 4

Else

ws.Cells(startingSummaryRow, 10).Interior.ColorIndex = 3

End If

'Find percent change

If YearBegOpenPrice = 0 Then
YearBegOpenPriceNonZero = 1
Else: YearBegOpenPriceNonZero = YearBegOpenPrice
End If

ws.Cells(startingSummaryRow, 11) = (YearEndClosePrice - YearBegOpenPrice) / YearBegOpenPriceNonZero

'formatting the Percent Change to %
ws.Range("K2:K" & lastRow).NumberFormat = "0.00%"

'increment starting summaryrow for the next ticker
startingSummaryRow = startingSummaryRow + 1

'reset the total charges
TotalStockVolume = 0

'check if the next row has same ticker, if yes, then sum the volume
Else

TotalStockVolume = TotalStockVolume + ws.Cells(i, 7).Value


End If

Next i

 'find the last row in summary table
 lastRowinSummary = ws.Cells(Rows.Count, "I").End(xlUp).Row

'Find Greatest Total Volume, percent Increase and percent decrease
    GreatestTotalVolume = WorksheetFunction.Max(ws.Range("L2:L" & lastRowinSummary))
    GreatestPercentIncrease = WorksheetFunction.Max(ws.Range("K2:K" & lastRowinSummary)) * 100
    GreatestPercentDecrease = WorksheetFunction.Min(ws.Range("K2:K" & lastRowinSummary)) * 100
  
GreatestTotalTicker = ""
GreatestPercentIncreaseTicker = ""
GreatestPercentDecreaseTicker = ""
 'loop through the summary table
         
         
    For j = 2 To lastRowinSummary
       If ws.Cells(j, 12) = GreatestTotalVolume Then
       GreatestTotalTicker = ws.Cells(j, 9)
             
       End If
       
       If ws.Cells(j, 11) * 100 = GreatestPercentIncrease Then
       GreatestPercentIncreaseTicker = ws.Cells(j, 9)
       
       End If
       
              
       If ws.Cells(j, 11) * 100 = GreatestPercentDecrease Then
       GreatestPercentDecreaseTicker = ws.Cells(j, 9)
       
       End If
             
    Next j
    
    'Display Greatest Total Volume, percent Increase and percent decrease
    
      'define the headers for greatest values
      
ws.Cells(1, 16).Value = "Ticker"
ws.Cells(1, 17).Value = "Value"
ws.Cells(2, 15).Value = "Greatest % Increase"
ws.Cells(3, 15).Value = "Greatest % Decrease"
ws.Cells(4, 15).Value = "Greatest Total Volume"

'print values in respective cells


       ws.Cells(2, 16).Value = GreatestPercentIncreaseTicker
       ws.Cells(3, 16).Value = GreatestPercentDecreaseTicker
       ws.Cells(4, 16).Value = GreatestTotalTicker
       
       ws.Cells(2, 17).Value = GreatestPercentIncrease / 100
       ws.Cells(2, 17).NumberFormat = "0.00%"
       ws.Cells(3, 17).Value = GreatestPercentDecrease / 100
       ws.Cells(3, 17).NumberFormat = "0.00%"
       ws.Cells(4, 17).Value = GreatestTotalVolume
       

Next

End Sub

