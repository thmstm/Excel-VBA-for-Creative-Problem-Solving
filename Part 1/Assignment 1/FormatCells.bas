Attribute VB_Name = "Module1"
Option Explicit
Sub FormatCells()
Attribute FormatCells.VB_ProcData.VB_Invoke_Func = " \n14"
'
' FormatCells Macro
'

' Formatting the active cell

    With Selection.Font
        .Name = "Arial"
        .Size = 11
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .ThemeColor = xlThemeColorLight1
        .TintAndShade = 0
        .ThemeFont = xlThemeFontNone
    End With
    With Selection.Font
        .Name = "Arial"
        .Size = 14
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .ThemeColor = xlThemeColorLight1
        .TintAndShade = 0
        .ThemeFont = xlThemeFontNone
    End With
    Selection.Font.Bold = True
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    
    ' Deleting the row below the active cell
    
    ActiveCell.Offset(1, 0).Rows("1:1").EntireRow.Select
    Selection.Delete Shift:=xlUp
    
    ' Deleting Column A
    
    Columns("A:A").Select
    Selection.Delete Shift:=xlToLeft
End Sub
