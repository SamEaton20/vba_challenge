Sub stock_market()
    Dim i As Integer
    Dim Ticker As String
    Dim Last_Row As Long
    Dim ws As Worksheet
    Dim Total_Stock_Volume As Double
    Dim Summary_Table As Integer
    Dim Open_Price As Double
    Dim Closing_Price As Double
    Dim Quarterly_Change As Double
    Dim Percent_Change As Double

    ' Loop through each worksheet in the workbook
    For Each ws In Worksheets
        ' Define Last row of worksheet
        Last_Row = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
        ' Create the column headings for summary table
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Quarterly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Stock Volume"
        
        ' Reset total stock volume for the worksheet
        Total_Stock_Volume = 0
        Open_Price = 0
        Closing_Price = 0
        Ticker = ""

        ' Variables to track greatest percentage changes and total volume for this worksheet
        Dim Greatest_Increase_Ticker As String
        Dim Greatest_Decrease_Ticker As String
        Dim Greatest_Volume_Ticker As String
        Dim Greatest_Increase As Double
        Dim Greatest_Decrease As Double
        Dim Greatest_Volume As Double

        ' Initialize variables for this worksheet
        Greatest_Increase = -1E+30 ' Very small initial value
        Greatest_Decrease = 1E+30 ' Very large initial value
        Greatest_Volume = 0
        Summary_Table = 2
        
        ' Loop through the rows of the current worksheet
        For i = 2 To Last_Row
            ' Check if it's a new ticker
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                ' Store the last closing price
                Closing_Price = ws.Cells(i, 6).Value
                
                ' Calculate quarterly change and percent change
                If Open_Price <> 0 Then
                    Quarterly_Change = Closing_Price - Open_Price
                    Percent_Change = (Quarterly_Change / Open_Price) * 100
                Else
                    Quarterly_Change = 0
                    Percent_Change = 0
                End If

                ' Record the ticker and total stock volume
                Ticker = ws.Cells(i, 1).Value
                ws.Range("I" & Summary_Table).Value = Ticker
                ws.Range("J" & Summary_Table).Value = Quarterly_Change
                ws.Range("K" & Summary_Table).Value = Percent_Change
                ws.Range("L" & Summary_Table).Value = Total_Stock_Volume
                
                ' Format Quarterly Change cell color based on value
                If Quarterly_Change < 0 Then
                    ws.Range("J" & Summary_Table).Interior.Color = RGB(255, 0, 0) ' Red
                ElseIf Quarterly_Change > 0 Then
                    ws.Range("J" & Summary_Table).Interior.Color = RGB(0, 255, 0) ' Green
                Else
                    ws.Range("J" & Summary_Table).Interior.ColorIndex = xlNone ' No color for zero
                End If

                ' Check for greatest % increase, decrease, and volume for this worksheet
                If Percent_Change > Greatest_Increase Then
                    Greatest_Increase = Percent_Change
                    Greatest_Increase_Ticker = Ticker
                End If
                
                If Percent_Change < Greatest_Decrease Then
                    Greatest_Decrease = Percent_Change
                    Greatest_Decrease_Ticker = Ticker
                End If
                
                If Total_Stock_Volume > Greatest_Volume Then
                    Greatest_Volume = Total_Stock_Volume
                    Greatest_Volume_Ticker = Ticker
                End If

                ' Increment summary table row
                Summary_Table = Summary_Table + 1
                
                ' Reset for the next ticker
                Total_Stock_Volume = 0
                Open_Price = 0
                
            Else
                ' Accumulate total stock volume
                Total_Stock_Volume = Total_Stock_Volume + ws.Cells(i, 7).Value
                
                ' Set Open_Price if it's the first row for this ticker
                If Open_Price = 0 Then
                    Open_Price = ws.Cells(i, 3).Value
                End If
            End If
            
        Next i

        ' Format percent change column as percentage
        ws.Range("K2:K" & Summary_Table - 1).NumberFormat = "0.00%"

        ' Output the greatest values for this worksheet
        ws.Range("M1").Value = "Greatest % Increase"
        ws.Range("N1").Value = "Greatest % Decrease"
        ws.Range("O1").Value = "Greatest Total Volume"

        ws.Range("M2").Value = Greatest_Increase_Ticker
        ws.Range("N2").Value = Greatest_Decrease_Ticker
        ws.Range("O2").Value = Greatest_Volume_Ticker
        ws.Range("M3").Value = Greatest_Increase & "%"
        ws.Range("N3").Value = Greatest_Decrease & "%"
        ws.Range("O3").Value = Greatest_Volume

    Next ws

    MsgBox "Stock market data processed successfully!"
End Sub
