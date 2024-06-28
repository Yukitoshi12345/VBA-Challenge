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
    For Each ws In worksheets
        
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
    

    Next ws

End Sub