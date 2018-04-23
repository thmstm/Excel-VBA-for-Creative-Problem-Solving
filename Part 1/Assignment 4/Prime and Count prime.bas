Attribute VB_Name = "Module1"
Option Explicit

Function prime(n As Integer) As Boolean
'Place your code here
Dim l As Integer, flag As Boolean, i As Integer
'Initializing
l = WorksheetFunction.RoundDown(Sqr(n), 0)
flag = True
For i = 2 To l
    If n Mod i = 0 Then flag = False
Next i
prime = flag
End Function

Function countprime(n1 As Integer, n2 As Integer) As Integer
'Place your code here
Dim i As Integer, count As Integer
For i = n1 To n2
    If prime(i) Then count = count + 1
Next i
countprime = count
End Function
