Sub StockAnalysis():
Dim ws As Worksheet
For Each ws In ThisWorkbook.Worksheets

'Declare the variables used
Dim i As Long
Dim Ticker As String
Dim Volume As LongLong
Dim TableRow As Integer
TableRow = 2
Dim OpenPrice As Double
Dim ClosePrice As Double
Dim YearlyChange As Double
Dim PercentChange As Double

'Set Table Headings
ws.Range("I1") = "Ticker"
ws.Range("J1") = "Yearly Change"
ws.Range("K1") = "Percent Change"
ws.Range("L1") = "Total Stock Volume"

'Set Summary Table Settings
ws.Range("O2") = "Greatest % Increase"
ws.Range("O3") = "Greatest % Decrease"
ws.Range("O4") = "Greatest Total Volume"

Dim LastRow As Long
LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

For i = 2 To LastRow

If ws.Cells(i, 2) = ws.Cells(2, 2).Value Then
OpenPrice = ws.Cells(i, 3).Value
'If the ticker is a new ticker
ElseIf ws.Cells(i, 1) <> ws.Cells(i + 1, 1) Then
Ticker = ws.Cells(i, 1).Value
ws.Range("I" & TableRow).Value = Ticker
ws.Range("L" & TableRow).Value = Volume

'Calculate the change in stock price
ClosePrice = ws.Cells(i, 6).Value
YearlyChange = OpenPrice - ClosePrice
ws.Range("J" & TableRow).Value = YearlyChange

'Change colour of cell for change in stock Price
If YearlyChange > 0 Then
ws.Range("J" & TableRow).Interior.ColorIndex = 4
ElseIf YearlyChange < 0 Then
ws.Range("J" & TableRow).Interior.ColorIndex = 3

End If

'Calculate the percentage change
PercentChange = YearlyChange / OpenPrice
ws.Range("K" & TableRow).Value = PercentChange
'Change format of Percent Change to Percent
ws.Range("K" & TableRow).Style = "Percent"

'Determine if the amount is the greatest change
If PercentChange > Range("P2").Value Then
ws.Range("P2").Value = PercentChange
ws.Range("Q2").Value = Ticker

'Determine if the amount is the smallest change
ElseIf PercentChange < Range("P3").Value Then
ws.Range("P3").Value = PercentChange
ws.Range("Q3").Value = Ticker

'Determine if the volume is the greatest
ElseIf Volume > ws.Range("P4").Value Then
ws.Range("P4").Value = Volume
ws.Range("Q4").Value = Ticker

End If
'Reset volume & add row to Summary Table
Volume = 0
TableRow = TableRow + 1
OpenPrice = 0
ClosePrice = 0

'If ticker is the same as above add the Volume
Else: Volume = Volume + ws.Cells(i, 7).Value

End If

Next i

'Space Columns accordingly
ws.Columns("I:L").AutoFit
ws.Columns("O:Q").AutoFit
ws.Range("P2").Style = "Percent"
ws.Range("P3").Style = "Percent"

Next ws

End Sub


