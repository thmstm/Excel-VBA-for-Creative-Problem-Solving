Attribute VB_Name = "Module1"
Option Explicit

Function tank(R As Double, H As Double, d As Double) As Double
'Place your code here
If d <= R Then
    tank = WorksheetFunction.Pi * d ^ 2 / 3 * (3 * R - d)
ElseIf d < (H - R) Then
    tank = 2 * WorksheetFunction.Pi * R ^ 3 / 3 + WorksheetFunction.Pi * R ^ 2 * (d - R)
ElseIf d <= H Then
    tank = 4 * WorksheetFunction.Pi * R ^ 3 / 3 + WorksheetFunction.Pi * R ^ 2 * (H - 2 * R) - WorksheetFunction.Pi * (H - d) ^ 2 / 3 * (3 * R - H + d)
Else
    tank = -1
End If
End Function
