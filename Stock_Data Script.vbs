Sub Stock_Data()
'Loop all sheets
Dim WS As Worksheet
For Each WS In ActiveWorkbook.Worksheets
WS.Activate
'last row
lastrow = WS.Cells(Rows.Count, 1).End(xlUp).Row
'add heading
Cells(1, "I") = "Ticker"
Cells(1, "J") = "Yearly Change"
Cells(1, "K") = "Percent Change"
Cells(1, "L") = "Total Stock Volume"
'varibales to hold value
Dim open_price As Double
Dim close_price As Double
Dim yearly_change As Double
Dim ticker_name As String
Dim percent_change As Double
Dim volume As Double
volume = 0
Dim r As Double
r = 2
Dim c As Double
c = 1
Dim i As Long
'set initial open price
open_price = Cells(2, c + 2)
'loop
For i = 2 To lastrow
If Cells(i + 1, c) <> Cells(i, c) Then
'ticker name
ticker_name = Cells(i, c)
Cells(r, c + 8) = ticker_name
'close price
close_price = Cells(i, c + 5)
'yearly change
yearly_change = close_price - open_price
Cells(r, c + 9) = yearly_change
'percent change
If (open_price = 0 And close_price = 0) Then
percent_change = 0
 ElseIf (open_price = 0 And close_price <> 0) Then
 percent_change = 1
Else
percent_change = yearly_change / open_price
Cells(r, c + 10) = percent_change
Cells(r, c + 10).NumberFormat = "0.00%"
End If
'total volume
Cells(r, c + 11) = volume + Cells(i, c + 6)
r = r + 1
open_price = Cells(i + 1, c + 2)
volume = 0
'if cells are the same ticker
Else
volume = volume + Cells(i, c + 6)
End If
Next i
 ' Determine the Last Row of Yearly Change per WS
        YCLastRow = WS.Cells(Rows.Count, c + 8).End(xlUp).Row
        ' Cell Colors
        For j = 2 To YCLastRow
            If (Cells(j, c + 9).Value > 0 Or Cells(j, c + 9).Value = 0) Then
                Cells(j, c + 9).Interior.ColorIndex = 10
            ElseIf Cells(j, c + 9).Value < 0 Then
                Cells(j, c + 9).Interior.ColorIndex = 3
            End If
        Next j
        Next
        End Sub
