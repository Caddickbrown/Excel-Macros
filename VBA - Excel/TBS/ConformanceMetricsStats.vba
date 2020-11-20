'Used to sort specific data into the correct columns'

Sub Conformance_MetricsSO_Stats()
'
' Conformance_MetricsSO_Stats Macro
'
'
    Application.ScreenUpdating = False
    Columns("L:L").Select
    Selection.Cut
    Columns("B:B").Select
    Selection.Insert Shift:=xlToRight
    Columns("M:O").Select
    Selection.Cut
    Columns("C:C").Select
    Selection.Insert Shift:=xlToRight
    Columns("Q:R").Select
    Selection.Cut
    Columns("F:F").Select
    Selection.Insert Shift:=xlToRight
    Columns("H:H").Select
    Columns("H:AJ").Select
    Selection.Delete Shift:=xlToLeft
    Cells.Select
    Cells.EntireColumn.AutoFit
    Range("A1").Select
    Application.ScreenUpdating = True
End Sub
