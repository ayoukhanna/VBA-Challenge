Attribute VB_Name = "Module2"
Sub Ticker()

Worksheets("2016").Select


Dim i As Long
Dim RowCount As Long
Dim Start As Long
Start = 2
Dim Total As Double
Total = 0
Dim change As Double
Dim opening As Double
Dim closing As Double
opening = Cells(2, 3).Value
Dim PctChange As Double

RowCount = Cells(Rows.Count, "A").End(xlUp).Row

For i = 2 To RowCount
Total = Total + Cells(i, 7).Value

If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then

Cells(Start, 9).Value = Cells(i, 1).Value
Cells(Start, 12).Value = Total
closing = Cells(i, 6).Value


change = closing - opening
opening = Cells(i + 1, 3).Value
Cells(Start, 10).Value = change
If opening > 0 Then
PctChange = change / opening
Else
PctChange = 0
End If

Cells(Start, 10).Value = change
If change > 0 Then
Cells(Start, 10).Interior.ColorIndex = 4
Else
Cells(Start, 10).Interior.ColorIndex = 3

End If

Start = Start + 1
Total = 0

End If

Next i
MsgBox ("Finished")

End Sub
