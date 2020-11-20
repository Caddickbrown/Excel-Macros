'Used to sort specific data into the correct columns'

Sub Conformance_MetricsSO()
'
' Conformance_MetricsSO Macro
'
'
    Application.ScreenUpdating = False
    Columns("G:G").Select
    Selection.Cut
    Columns("E:E").Select
    Selection.Insert Shift:=xlToRight
    Columns("H:H").Select
    Selection.Cut
    Columns("F:F").Select
    Selection.Insert Shift:=xlToRight
    Columns("I:K").Select
    Selection.Cut
    Columns("G:G").Select
    Selection.Insert Shift:=xlToRight
    Columns("T:T").Select
    Selection.Cut
    Columns("J:J").Select
    Selection.Insert Shift:=xlToRight
    Columns("K:K").Select
    Columns("BM:BM").Select
    Selection.Cut
    Columns("K:K").Select
    Selection.Insert Shift:=xlToRight
    Columns("L:L").Select
    Columns("L:CG").Select
    Selection.Delete Shift:=xlToLeft
    Cells.Select
    Cells.EntireColumn.AutoFit
    Range("A1").Select
    Application.ScreenUpdating = True
End Sub
