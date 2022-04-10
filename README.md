



Option Explicit

Sub vba_challenge()
    Const VOL_COL As Integer = 7
    Const TICKER_COL As Integer = 1
    Const OPENING_COL As Integer = 3
    
    Dim SummaryRow As Long
    Dim currentrow As Long
    Dim totalvolume As Double
    Dim ws As Worksheet
    Dim openingprice As Double
    Dim closingprice As Double
    Dim yearlychange As Double
    Dim currentticker As String
    Dim LastRow As Long
    Dim percentchange As Double
    Dim maxperincrease As Double
    Dim maxperticker As String
    Dim minperincrease As Double
    Dim minperticker As String
    Dim greatestvolume As Double
    Dim greatestvolumeticker As String
    
    'loop through all the worksheets
    'your going to want the ticker in  a summary table
    For Each ws In Worksheets
    
        SummaryRow = 2
      
        maxperincrease = -9999
        maxperticker = ""
        minperincrease = 9999
        minperticker = ""
        greatestvolume = 0
        
        'find the last row subtract one to return the number of rows without the header
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        openingprice = ws.Cells(2, 3).Value
        'loop through all rows (use loooping variable i)
        For currentrow = 2 To LastRow
            If ws.Cells(currentrow, TICKER_COL).Value <> ws.Cells(currentrow - 1, TICKER_COL).Value Then
                'this is the first row; save the current ticker, set volume to zero, save the opening price
                currentticker = ws.Cells(currentrow, TICKER_COL).Value
                totalvolume = 0
                openingprice = ws.Cells(currentrow, OPENING_COL).Value
            End If
            
            'for every row count the volume (including first and last)
            totalvolume = totalvolume + ws.Cells(currentrow, VOL_COL).Value
         
            'is this the last row of the ticker
            If ws.Cells(currentrow, TICKER_COL).Value <> ws.Cells(currentrow + 1, TICKER_COL).Value Then
                    'assign closing price
                    closingprice = ws.Cells(currentrow, 6).Value
                    
                    'Calculatations (per stock)
                    yearlychange = closingprice - openingprice
                   
                    If (openingprice > 0) Then
                        percentchange = yearlychange / openingprice
                    Else
                        percentchange = 0
                        MsgBox ("The following ticker opening price is zero " & currentticker)
                    End If
                    
                    'Calculations (per sheet)
                    If percentchange > maxperincrease Then
                        maxperincrease = percentchange
                        maxperticker = currentticker
                    End If
                    If percentchange < minperincrease Then
                        minperincrease = percentchange
                        minperticker = currentticker
                    End If
                    If totalvolume > greatestvolume Then
                        greatestvolume = totalvolume
                        greatestvolumeticker = currentticker
                    End If
                    
                    'output section each stock
                    ws.Cells(SummaryRow, 9).Value = currentticker
                    ws.Cells(SummaryRow, 10).Value = yearlychange
                    'color yearln based off positive or negative changey change red or green
                    If (yearlychange > 0) Then
                        ws.Cells(SummaryRow, 10).Interior.ColorIndex = 4
                    Else
                        ws.Cells(SummaryRow, 10).Interior.ColorIndex = 3
                    End If
                    ws.Cells(SummaryRow, 11).Value = percentchange
                    ws.Cells(SummaryRow, 12).Value = totalvolume
                    'prepare for the next stock
                    SummaryRow = SummaryRow + 1
            End If
        Next currentrow
        
        'Output section each sheet (Bonus Section)
        ws.Cells(1, 16).Value = "Ticker"
        ws.Cells(1, 17).Value = "Value"
        ws.Cells(2, 15).Value = "Greatest % Increase"
        ws.Cells(2, 16).Value = maxperticker
        ws.Cells(2, 17).Value = maxperincrease
        ws.Cells(2, 17).NumberFormat = "0.00%"
        ws.Cells(3, 15).Value = "Greatest % Decrease"
        ws.Cells(3, 16).Value = minperticker
        ws.Cells(3, 17).Value = minperincrease
        ws.Cells(3, 17).NumberFormat = "0.00%"
        ws.Cells(4, 15).Value = "Greatest Total Volume"
        ws.Cells(4, 16).Value = greatestvolumeticker
        ws.Cells(4, 17).Value = greatestvolume
        ws.Cells(4, 17).NumberFormat = "General"
    Next ws

End Sub

   
