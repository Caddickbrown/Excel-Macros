

Sub SDL04_Priorities()
'
' SDL04_Priorities Macro
'

'
    Application.ScreenUpdating = False
    Columns("H:H").Select
    Selection.Cut
    Columns("C:C").Select
    Selection.Insert Shift:=xlToRight
    Columns("J:J").Select
    Selection.Cut
    Columns("D:D").Select
    Selection.Insert Shift:=xlToRight
    Columns("I:I").Select
    Selection.Cut
    Columns("E:E").Select
    Selection.Insert Shift:=xlToRight
    Columns("F:F").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.ClearContents
    Cells.Select
    Cells.EntireColumn.AutoFit
    Range("A1").Select
    Application.ScreenUpdating = True
End Sub
