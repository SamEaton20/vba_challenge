Sub stock_market()
Dim i As Integer
Dim Ticker As String
Dim Ticker_Total As Double
    Ticker_Total = 0
Dim Last_row As Long
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
Dim Percentage_Change As Double
    
'Loop through each worksheet in the workbook
'For Each ws In Worksheets

'Define Lastrow of worksheet
Last_row = Cells(Rows.Count, 1).End(xlUp).Row

'Create the column headings for summary table
Range("I1").Value = "Ticker"
Range("J1").Value = "Quarterly Change"
Range("K1").Value = "Percent Change"
Range("L1").Value = "Total Stock Volume"

Range("O1").Value = "Ticker"
Range("P1").Value = "Value"
Range("N2").Value = "Greatest % Increase"
Range("N3").Value = "Greatest % Decrease"
Range("N4").Value = "Greatest Total Volume"

'Do loop of current worksheet to Lastrow
For i = 2 To Last_row

'See the Ticker total
If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
    Ticker = Cells(i, 1).Value
    Ticker_Total = Ticker_Total + Cells(i, 7).Value
    Range("I" & Summary_Table).Value = Ticker
    Range("L" & Summary_Table).Value = Total_Stock_Volume
    Summary_Table = Summary_Table + 1
    Total_Stock_Volume = 0
    First_Row = 0
Else
    Total_Stock_Volume = Total_Stock_Volume + Cells(i, 7).Value
    If First_Row = 0 Then
        Open_Price = Cells(i, 3).Value
        First_Row = 1
        Closing_Price = Cells(i, 6).Value
        Last_row = 1
    End If
'Getting the Quarterly Change


End If

Next i

'Next ws

End Sub
