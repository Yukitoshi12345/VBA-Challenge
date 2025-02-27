Attribute VB_Name = "StockAnalysis"

' ---------------------------------------------------------------------------------------------------------------------------------------------
' Create a script that loops through all the stocks for each quarter and outputs the following information:

' - The ticker symbol
' - Quarterly change from the opening price at the beginning of a given quarter to the closing price at the end of that quarter.
' - The percentage change from the opening price at the beginning of a given quarter to the closing price at the end of that quarter.
' - The total stock volume of the stock.
' - Add functionality to your script to return the stock with the "Greatest % increase", "Greatest % decrease", and "Greatest total volume".
' - Make the appropriate adjustments to your VBA script to enable it to run on every worksheet (that is, every quarter) at once.
' ---------------------------------------------------------------------------------------------------------------------------------------------

Sub StockAnalysis()

    ' Loop through each worksheet
    For Each ws In Worksheets
    
        ' Declare variables
        Dim ticker As String
        Dim openPrice As Double
        Dim closePrice As Double
        Dim quarterlyChange As Double
        Dim percentChange As Double
        Dim greatestInc As Double
        Dim greatestDec As Double
        Dim greatestVol As Double
        Dim totalVolume As Double
        Dim summaryTableRow As Integer
        Dim lastRow As Long
        Dim lastRowSummary As Long
        
        ' Set Column Headers
        With ws
            ws.Cells(1, 9).Value = "Ticker"
            ws.Cells(1, 10).Value = "Quarterly Change"
            ws.Cells(1, 11).Value = "Percent Change"
            ws.Cells(1, 12).Value = "Total Stock Volume"
            
            ws.Cells(1, 16).Value = "Ticker"
            ws.Cells(1, 17).Value = "Value"
            
            ws.Cells(2, 15).Value = "Greatest % Increase"
            ws.Cells(3, 15).Value = "Greatest % Decrease"
            ws.Cells(4, 15).Value = "Greatest Total Volume"
        End With
            
        ' Autofit Columns
        With ws
            ws.Columns(2).AutoFit
            ws.Columns(9).AutoFit
            ws.Columns(10).AutoFit
            ws.Columns(11).AutoFit
            ws.Columns(12).AutoFit
            ws.Columns(15).AutoFit
            ws.Columns(16).AutoFit
            ws.Columns(17).AutoFit
        End With
    
        ' Initialise variables
        summaryTableRow = 2
        totalVolume = 0
        
        ' Find the last row with data in the worksheet
        lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        
        ' Initialise open price with the first ticker's open price
        openPrice = ws.Cells(2, 3).Value
        
        ' Loop through all rows of data
        For i = 2 To lastRow
            
            ' Check if the next row has a different ticker symbol
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                
                ' Capture the current ticker symbol
                ticker = ws.Cells(i, 1).Value
                
                ' Capture the closing price of the current ticker
                closePrice = ws.Cells(i, 6).Value
                
                ' Calculate quarterly change and percent change
                quarterlyChange = closePrice - openPrice
                If openPrice <> 0 Then
                    percentChange = quarterlyChange / openPrice
                Else
                    percentChange = 0
                End If
                
                ' Output results to the summary table
                ws.Cells(summaryTableRow, 9).Value = ticker
                ws.Cells(summaryTableRow, 10).Value = quarterlyChange
                ws.Cells(summaryTableRow, 11).Value = percentChange
                ws.Cells(summaryTableRow, 11).NumberFormat = "0.00%"
                ws.Cells(summaryTableRow, 12).Value = totalVolume
                
                ' Format the quarterly change cells
                If quarterlyChange > 0 Then
                    ws.Cells(summaryTableRow, 10).Interior.ColorIndex = 4 ' Green for positive change
                    ws.Cells(summaryTableRow, 11).Interior.ColorIndex = 4 ' Green for positive change
                ElseIf quarterlyChange < 0 Then
                    ws.Cells(summaryTableRow, 10).Interior.ColorIndex = 3 ' Red for negative change
                    ws.Cells(summaryTableRow, 11).Interior.ColorIndex = 3 ' Red for negative change
                End If
                
                ' Prepare for the next ticker
                openPrice = ws.Cells(i + 1, 3).Value
                totalVolume = 0
                summaryTableRow = summaryTableRow + 1
                
            Else
                ' Accumulate the total volume for the current ticker
                totalVolume = totalVolume + ws.Cells(i, 7).Value
            End If
            
        Next i
        
        ' Initialise greatest increase, decrease, and volume variables
        greatestInc = ws.Cells(2, 11).Value
        greatestDec = ws.Cells(2, 11).Value
        greatestVol = ws.Cells(2, 12).Value
        
        ' Find the last row in the summary table
        lastRowSummary = ws.Cells(ws.Rows.Count, 9).End(xlUp).Row
        
        ' Loop through the summary table to find the greatest increase, decrease, and volume
        For i = 2 To lastRowSummary
            If ws.Cells(i, 11).Value > greatestInc Then
                greatestInc = ws.Cells(i, 11).Value
                tickerInc = ws.Cells(i, 9).Value
            End If
            
            If ws.Cells(i, 11).Value < greatestDec Then
                greatestDec = ws.Cells(i, 11).Value
                tickerDec = ws.Cells(i, 9).Value
            End If
            
            If ws.Cells(i, 12).Value > greatestVol Then
                greatestVol = ws.Cells(i, 12).Value
                tickerVol = ws.Cells(i, 9).Value
            End If
        Next i
        
        ' Output the greatest increase, decrease, and volume to the secondary summary table
        ws.Cells(2, 16).Value = tickerInc
        ws.Cells(2, 17).Value = greatestInc
        ws.Cells(2, 17).NumberFormat = "0.00%"
        
        ws.Cells(3, 16).Value = tickerDec
        ws.Cells(3, 17).Value = greatestDec
        ws.Cells(3, 17).NumberFormat = "0.00%"
        
        ws.Cells(4, 16).Value = tickerVol
        ws.Cells(4, 17).Value = greatestVol
    
    Next ws
    
End Sub