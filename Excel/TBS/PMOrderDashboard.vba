

Sub PM_OrderDashboard()
'
' PM_OrderDashboard Macro
'
'
    Application.ScreenUpdating = False
    Columns("G:G").Select
    Selection.Cut
    Columns("E:E").Select
    Selection.Insert Shift:=xlToRight
    Columns("I:I").Select
    Selection.Cut
    Columns("F:F").Select
    Selection.Insert Shift:=xlToRight
    Columns("J:J").Select
    Selection.Cut
    Columns("G:G").Select
    Selection.Insert Shift:=xlToRight
    Columns("J:J").Select
    Selection.Cut
    Columns("H:H").Select
    Selection.Insert Shift:=xlToRight
    Columns("I:I").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Columns("L:L").Select
    Selection.Cut
    Columns("J:J").Select
    Selection.Insert Shift:=xlToRight
    Columns("K:K").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.ClearContents
    Selection.End(xlToLeft).Select
    Selection.End(xlToLeft).Select
    Selection.End(xlToLeft).Select
    Cells.Select
    Cells.EntireColumn.AutoFit
    Range("A1").Select
    Application.ScreenUpdating = True
End Sub
