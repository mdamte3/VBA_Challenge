Attribute VB_Name = "Module1"
Sub StockAnalysis()
    ' Declare variables for tracking stock information
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim ticker As String
    Dim openingPrice As Double
    Dim closingPrice As Double
    Dim yearlyChange As Double
    Dim percentChange As Double
    Dim totalVolume As Double
    Dim summaryRow As Integer
 
    ' Variables for tracking greatest values
    Dim greatestIncrease As Double
    Dim greatestDecrease As Double
    Dim greatestVolume As Double
    Dim increaseTickerName As String
    Dim decreaseTickerName As String
    Dim volumeTickerName As String
 
    ' Loop through each worksheet in the workbook
    For Each ws In ThisWorkbook.Worksheets
        ' Initialize variables for each worksheet
        lastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        summaryRow = 2
        totalVolume = 0
        greatestIncrease = 0
        greatestDecrease = 0
        greatestVolume = 0
 
        ' Create headers for summary table
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Yearly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Stock Volume"
 
        ' Get the opening price for the first ticker
        openingPrice = ws.Cells(2, 3).Value
 
        ' Loop through all rows of data
        For i = 2 To lastRow
            ' Add to total volume
            totalVolume = totalVolume + ws.Cells(i, 7).Value
 
            ' Check if we're still within the same ticker
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                ' Set the ticker symbol
                ticker = ws.Cells(i, 1).Value
 
                ' Get the closing price
                closingPrice = ws.Cells(i, 6).Value
 
                ' Calculate yearly change
                yearlyChange = closingPrice - openingPrice
 
                ' Calculate percent change
                If openingPrice <> 0 Then
                    percentChange = yearlyChange / openingPrice
                Else
                    percentChange = 0
                End If
 
                ' Write the values to the summary table
                ws.Cells(summaryRow, 9).Value = ticker
                ws.Cells(summaryRow, 10).Value = yearlyChange
                ws.Cells(summaryRow, 11).Value = percentChange
                ws.Cells(summaryRow, 12).Value = totalVolume
 
                ' Format the yearly change cell
                If yearlyChange >= 0 Then
                    ws.Cells(summaryRow, 10).Interior.Color = RGB(0, 255, 0)  ' Green
                Else
                    ws.Cells(summaryRow, 10).Interior.Color = RGB(255, 0, 0)  ' Red
                End If
 
                ' Format percent change as percentage
                ws.Cells(summaryRow, 11).NumberFormat = "0.00%"
 
                ' Check for greatest values
                If percentChange > greatestIncrease Then
                    greatestIncrease = percentChange
                    increaseTickerName = ticker
                End If
 
                If percentChange < greatestDecrease Then
                    greatestDecrease = percentChange
                    decreaseTickerName = ticker
                End If
 
                If totalVolume > greatestVolume Then
                    greatestVolume = totalVolume
                    volumeTickerName = ticker
                End If
 
                ' Reset variables for next ticker
                summaryRow = summaryRow + 1
                totalVolume = 0
                openingPrice = ws.Cells(i + 1, 3).Value
            End If
        Next i
 
        ' Create greatest values summary
        ws.Range("O2").Value = "Greatest % Increase"
        ws.Range("O3").Value = "Greatest % Decrease"
        ws.Range("O4").Value = "Greatest Total Volume"
        ws.Range("P1").Value = "Ticker"
        ws.Range("Q1").Value = "Value"
 
        ' Write greatest values
        ws.Range("P2").Value = increaseTickerName
        ws.Range("P3").Value = decreaseTickerName
        ws.Range("P4").Value = volumeTickerName
        ws.Range("Q2").Value = greatestIncrease
        ws.Range("Q3").Value = decreaseTickerName
        ws.Range("Q4").Value = greatestVolume
 
        ' Format greatest percentage cells
        ws.Range("Q2:Q3").NumberFormat = "0.00%"
 
        ' Auto-fit columns
        ws.Columns("I:Q").AutoFit
    Next ws
 
    ' Display completion message
    MsgBox "Stock analysis complete!"
End Sub

