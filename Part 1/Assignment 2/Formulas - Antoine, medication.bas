Attribute VB_Name = "Module1"
Option Explicit

Function antoine(A As Double, B As Double, C As Double, t As Double) As Double
'Place your code here
antoine = 10 ^ (A - (B / (t + C)))
End Function

Function medication(C0 As Double, k As Double, t As Double) As Double
'Place your code here
medication = C0 * Exp(-k * t)
End Function
