Attribute VB_Name = "Module1"
Option Explicit

Sub HighTemp()
Dim nr As Integer, i As Integer, HT As Double, c As Integer
Range("A2").Select
Call Reset
Application.ScreenUpdating = False
Selection.CurrentRegion.Select
HT = InputBox("Display the days in the last year that exceeded what temperature?")
nr = Selection.Rows.Count
For i = 2 To nr
    If Selection.Cells(i, 4) > HT Then
        c = c + 1
        If c = 1 Then
            Range("G1") = "Date"
            Range("H1") = "Temperature"
        End If
        Range("G" & c + 1) = Selection.Cells(i, 2) & "/" & Selection.Cells(i, 3) & "/" & Selection.Cells(i, 1)
        Range("H" & c + 1) = Selection.Cells(i, 4)
    End If
Next i
Range("A1").Select
End Sub

Sub LowTemp()
Dim nr As Integer, i As Integer, LT As Double, c As Integer
Call Reset
Application.ScreenUpdating = False
Selection.CurrentRegion.Select
LT = InputBox("Display the days in the last year that were below what temperature?")
nr = Selection.Rows.Count
For i = 2 To nr
    If Selection.Cells(i, 5) < LT Then
        c = c + 1
        If c = 1 Then
            Range("G1") = "Date"
            Range("H1") = "Temperature"
        End If
        Range("G" & c + 1) = Selection.Cells(i, 2) & "/" & Selection.Cells(i, 3) & "/" & Selection.Cells(i, 1)
        Range("H" & c + 1) = Selection.Cells(i, 5)
    End If
Next i
Range("A1").Select
End Sub
Sub Reset()
Columns("G:H").Clear
End Sub
