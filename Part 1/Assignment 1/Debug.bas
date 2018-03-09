Attribute VB_Name = "Module1"
Option Explicit

Sub Debugging()
Dim i As Integer, j As Integer, k As Integer, p As Double, s As Double
For i = 1 To 20
    For j = 1 To 20
        s = s + 0.03 * i + 0.07 * j
    Next j
Next i
p = s
For k = 1 To 2000
    p = p + 0.0005 * k ^ 2
    Debug.Assert p < 200000
Next k
MsgBox ("The final result is " & k)
End Sub



