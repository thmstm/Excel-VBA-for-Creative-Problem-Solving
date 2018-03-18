Attribute VB_Name = "Module1"
Option Explicit

Sub AddNumbersA()
'Place your code here
Range("G12") = Range("D4") + InputBox("Enter a number to add to the value of cell D4:")
End Sub

Sub AddNumbersB()
'Place your code here
ActiveCell.Offset(-3, 2) = ActiveCell + InputBox("Enter a number to add to the active cell:")
End Sub

Sub WherePutMe()
'Place your code here
Range(InputBox("Enter the letter of the column of target cell:") & InputBox("Enter the number of the row of the target cell:")) = Selection.Cells(2, 2)
End Sub

Sub Swap()
'Place your code here
Dim x As Double
x = Selection.Cells(1, 1)
Selection.Cells(1, 1) = Selection.Cells(1, 2)
Selection.Cells(1, 2) = x
End Sub
