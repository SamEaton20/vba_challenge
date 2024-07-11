Sub stock_market()
Dim i As Integer
Dim Ticker As String
Dim Ticker_Total As Double
    Ticker_Total = 0
Dim Last_Row As Long
Dim ws As Worksheet
Dim Total_Stock_Volume As Double
    Total_Stock_Volume = 0
Dim Summary_Table As Integer
    Summary_Table = 2
Dim First_Row As Integer
    First_Row = 0
Dim Open_Price As String
Dim Closing_Price As String
Dim Quarterly_Change As Double
Dim Percent_Change As Double
    
'Loop through each worksheet in the workbook
For Each ws In Worksheets

'Define Lastrow of worksheet
Last_Row = ws.Cells(Rows.Count, 1).End(xlUp).Row

'Create the column headings for summary table
ws.Range("I1").Value = "Ticker"
ws.Range("J1").Value = "Quarterly Change"
ws.Range("K1").Value = "Percent Change"
ws.Range("L1").Value = "Total Stock Volume"

ws.Range("O1").Value = "Ticker2"
ws.Range("P1").Value = "Value"
ws.Range("N2").Value = "Greatest % Increase"
ws.Range("N3").Value = "Greatest % Decrease"
ws.Range("N4").Value = "Greatest Total Volume"

'Do loop of current worksheet to Lastrow
For i = 2 To Last_Row

'See the Ticker total
If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
    Ticker = ws.Cells(i, 1).Value
    Ticker_Total = Ticker_Total + ws.Cells(i, 7).Value
    ws.Range("I" & Summary_Table).Value = Ticker
    ws.Range("L" & Summary_Table).Value = Total_Stock_Volume
    Summary_Table = Summary_Table + 1
    Total_Stock_Volume = 0
    First_Row = 0
Else
    Total_Stock_Volume = Total_Stock_Volume + ws.Cells(i, 7).Value
      If First_Row = 0 Then
        Open_Price = ws.Cells(i, 3).Value
        First_Row = 1
        Closing_Price = ws.Cells(i, 6).Value
        Last_Row = 0
End If
    
'Getting the Quarterly Change
If Open_Price <> 0 Then
    Quarterly_Change = (Closing_Price - Open_Price)
    Range("J" & Summary_Table).Value = Quarterly_Change
Else
    Quarterly_Change = 0
End If

'Getting the Percent Change
If Open_Price <> 0 Then
    Percent_Change = (Quarterly_Change / Open_Price) * 100
    Range("K" & Summary_Table).Value = Percent_Change
End If

    
End If
    
Next i

Next ws

End Sub

