Attribute VB_Name = "Module1"
Sub stockModule()

'navigators/utility variables
Dim summaryRowVar As Long
Dim summaryColVar As Long
Dim bonusSummaryColVar As Long
Dim sheetLastRowVar As Long

'values to find
Dim tickerVar As String
Dim openPriceVar As Double
Dim closePriceVar As Double
Dim stockTotalVar As LongLong
Dim percentChangeVar As Double
Dim maxVar As Double
Dim minVar As Double
Dim maxVolume As LongLong
Dim summaryVar As Integer

'move through each worksheet with this code
For Each ws In Worksheets

'find the last row of relevant data for the iterator range
sheetLastRowVar = ws.Cells(Rows.Count, "A").End(xlUp).Row

'find the very first open price for the sheet
openPriceVar = ws.Cells(2, 3).Value

'create the summary table
    'enter in headers
summaryColVar = ws.Cells(1, Columns.Count).End(xlToLeft).Column + 2
ws.Cells(1, summaryColVar) = "Ticker"
ws.Cells(1, summaryColVar + 1) = "Yearly Change"
ws.Cells(1, summaryColVar + 2) = "Percent Change"
ws.Cells(1, summaryColVar + 3) = "Total Stock Volume"

    'set summary row position
summaryRowVar = 2
    
    'set stock total
stockTotalVar = 0

For i = 2 To sheetLastRowVar
    'look for unique ticker values
If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then
    'on unique ticker value, build that ticker's row on the summary table
        'put in ticker name
        ws.Cells(summaryRowVar, summaryColVar) = ws.Cells(i, 1).Value
        'get the yearly change
            'get the year's closing price
        closePriceVar = ws.Cells(i, 6).Value
            'enter yearly change
        ws.Cells(summaryRowVar, summaryColVar + 1) = closePriceVar - openPriceVar
        'get the percent change
            'use a conditional to deal with open prices of 0. If both the open and close price _
            are 0, make  the percentage change 0. Otherwise, enter into the cell "Open price at 0"
        If openPriceVar = 0 Then
            If openPriceVar = closePriceVar Then
            percentChangeVar = 0
            Else
            ws.Cells(summaryRowVar, summaryColVar + 2) = "Open Price at 0"
            End If
        Else
            percentChangeVar = (closePriceVar - openPriceVar) / openPriceVar
        End If
        ws.Cells(summaryRowVar, summaryColVar + 2) = percentChangeVar
        'get the total volume after adding current volume to stock total
        ws.Cells(summaryRowVar, summaryColVar + 3) = stockTotalVar + ws.Cells(i, 7).Value
    'reset total volume
        stockTotalVar = 0
    'get the open price for new ticker
        openPriceVar = ws.Cells(i + 1, 3)
    'move down a row on the summary table for new entry
        summaryRowVar = summaryRowVar + 1
    Else
    'add current volume to stock total
    stockTotalVar = stockTotalVar + ws.Cells(i, 7).Value
End If
Next i


'format summary table
    'if positive=green, if negative=red, if 0= no change
    For j = 2 To summaryRowVar
        If ws.Range("J" & j) > 0 Then
            ws.Range("J" & j).Interior.ColorIndex = 4
        ElseIf ws.Range("J" & j) < 0 Then
           ws.Range("J" & j).Interior.ColorIndex = 3
        End If
    Next j
'change formatting for percentages
    ws.Range("K2:K" & summaryRowVar).NumberFormat = "0.00%"
    

    
'create the bonus summary table
'find start position
bonusSummaryColVar = ws.Cells(1, Columns.Count).End(xlToLeft).Column + 3
ws.Cells(1, bonusSummaryColVar + 1) = "Ticker"
ws.Cells(2, bonusSummaryColVar) = "Greatest % Increase"
ws.Cells(3, bonusSummaryColVar) = "Greatest % Decrease"
ws.Cells(4, bonusSummaryColVar) = "Greatest Total Volume"

'find the greatest percent increase
maxVar = WorksheetFunction.Max(ws.Range("K:K"))
'get the ticker
    'find the row of the value
summaryVar = WorksheetFunction.Match(maxVar, ws.Range("K:K"), 0)
    'get ticker put the ticker into the summary
ws.Cells(2, bonusSummaryColVar + 1) = ws.Cells(summaryVar, 9).Value
    'put the value into the summary
ws.Cells(2, bonusSummaryColVar + 2) = maxVar

'find the greatest % decrease
minVar = WorksheetFunction.Min(ws.Range("K:K"))

    'find the row of the value, and
summaryVar = WorksheetFunction.Match(minVar, ws.Range("K:K"), 0)
    'get the ticker for that value and put the ticker into the summary
ws.Cells(3, bonusSummaryColVar + 1) = ws.Cells(summaryVar, 9).Value
    'put the value into the summary
ws.Cells(3, bonusSummaryColVar + 2) = minVar

'greatest total volume
maxVolume = WorksheetFunction.Max(ws.Range("L:L"))
    'find the row of the value
summaryVar = WorksheetFunction.Match(maxVolume, ws.Range("L:L"), 0)
    'get the ticker for that value and put the ticker into the summary
ws.Cells(4, bonusSummaryColVar + 1) = ws.Cells(summaryVar, 9).Value
    'put the value into the summary
ws.Cells(4, bonusSummaryColVar + 2) = maxVolume
'format percentages for bonustable
 ws.Range("Q2", "Q3").NumberFormat = "0.00%"


Next ws

End Sub


