Attribute VB_Name = "CompileStockData"
Sub CompileStockData()

' Declare Variables
Dim r As Long
Dim ticker As String
Dim volume As Long
Dim rowStart As Long
Dim rowSummary As Long
Dim colYearlyChange As Range
Dim colPercentChange As Range
Dim colTotalStock As Range
Dim bestPercent As Double
Dim bestPercentTicker As String
Dim worstPercent As Double
Dim worstPercentTicker As String
Dim bestTotalStock As LongLong ' **LongLong? who woulda thunk: https://docs.microsoft.com/en-us/office/vba/language/reference/user-interface-help/longlong-data-type
Dim bestTotalStockTicker As String

' Loop all worksheets
For Each ws In Worksheets

    ' Format Header Cells and Percent columns
    ws.Range("K:K").NumberFormat = "0.00%"
    ws.Range("Q2:Q3").NumberFormat = "0.00%"
    ws.Range("Q4").NumberFormat = "#,##0"
    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Yearly Change"
    ws.Cells(1, 11).Value = "Percent Change"
    ws.Cells(1, 12).Value = "Total Stock Volume"
    ws.Cells(2, 15).Value = "Greatest % Increase"
    ws.Cells(3, 15).Value = "Greatest % Decrease"
    ws.Cells(4, 15).Value = "Greatest Total Volume"
    ws.Cells(1, 16).Value = "Ticker"
    ws.Cells(1, 17).Value = "Value"
    
    ' Assign Variables with initial ticker values
    r = 2
    ticker = ws.Cells(2, 1).Value
    rowStart = 2
    rowSummary = 2
    bestPercent = 0
    worstPercent = 0
    bestTotalStock = 0
    
    
    
    ' While there is a value in the ticker column, run the loop
    While ws.Cells(r, 1).Value <> ""
    
        ' While the next ticker is the same as the current, increment
        ' then, display the previous ticker, increment, and initalize new ticker
        While ticker = ws.Cells(r + 1, 1).Value
            r = r + 1
        Wend
        
        ' Display the previous ticker info:
        ' Ticker, Yearly Change, Percent Change, and Total Stock Volume
        ws.Cells(rowSummary, 9).Value = ticker
        ws.Cells(rowSummary, 10).Value = ws.Cells(r, 6).Value - ws.Cells(rowStart, 3).Value
        
        ' Check for possible divide by 0 error for Percent Change
        If ws.Cells(rowStart, 3).Value > 0 Then
            ws.Cells(rowSummary, 11).Value = ws.Cells(rowSummary, 10).Value / ws.Cells(rowStart, 3).Value
        Else
            ws.Cells(rowSummary, 11).Value = 0
        End If
        
        ' Bonus: Check if above is greatest % increase or decrease
        ' First time through, it is both so assign to both
        If rowSummary = 2 Then
            bestPercent = ws.Cells(rowSummary, 11).Value
            bestPercentTicker = ws.Cells(rowSummary, 9).Value
            worstPercent = ws.Cells(rowSummary, 11).Value
            worstPercentTicker = ws.Cells(rowSummary, 9).Value
        ElseIf ws.Cells(rowSummary, 11).Value > bestPercent Then
            bestPercent = ws.Cells(rowSummary, 11).Value
            bestPercentTicker = ws.Cells(rowSummary, 9).Value
        ElseIf ws.Cells(rowSummary, 11).Value < worstPercent Then
            worstPercent = ws.Cells(rowSummary, 11).Value
            worstPercentTicker = ws.Cells(rowSummary, 9).Value
        End If
           
        ' Total Stocks Sum
        ' Reference used - https://stackoverflow.com/questions/11707888/sum-function-in-vba
        ws.Cells(rowSummary, 12).Formula = "=SUM(" & ws.Range(ws.Cells(rowStart, 7), ws.Cells(r, 7)).Address(False, False) & ")"
            
        ' Bonus: Check if above is greatest total volume
        If ws.Cells(rowSummary, 12).Value > bestTotalStock Then
            bestTotalStock = ws.Cells(rowSummary, 12).Value
            bestTotalStockTicker = ws.Cells(rowSummary, 9).Value
        End If
        
        ' Increment and initialize new ticker
        rowSummary = rowSummary + 1
        r = r + 1
        rowStart = r
        ticker = ws.Cells(r, 1).Value
        volume = ws.Cells(r, 7).Value
    Wend
    
    ' Conditional Formatting for the Yearly Change column
    ' **Reference used - https://www.automateexcel.com/vba/conditional-formatting/#A_Simple_Example_of_Creating_a_Conditional_Format_on_a_Range
    ' **Reference used - https://docs.microsoft.com/en-us/office/vba/api/excel.formatconditions.add
    
    ' First, capture the range and make them red by default
    Set colYearlyChange = ws.Range(ws.Cells(2, 10), ws.Cells(rowSummary - 1, 10))
    colYearlyChange.Interior.Color = RGB(254, 3, 2)
    
    ' then, make a cell green if >= 0
    colYearlyChange.FormatConditions.Add Type:=xlCellValue, Operator:=xlGreaterEqual, _
        Formula1:="=0"
    colYearlyChange.FormatConditions(1).Interior.Color = RGB(22, 224, 35)
    
    ' Bonus: Display superlative values
    ws.Range("P2").Value = bestPercentTicker
    ws.Range("Q2").Value = bestPercent
    ws.Range("P3").Value = worstPercentTicker
    ws.Range("Q3").Value = worstPercent
    ws.Range("P4").Value = bestTotalStockTicker
    ws.Range("Q4").Value = bestTotalStock

    ' Format the width of the columns to avoid overflow
    ' Reference used - https://docs.microsoft.com/en-us/office/vba/api/excel.range.autofit
    ws.Columns("I:Q").AutoFit

Next ws

End Sub
