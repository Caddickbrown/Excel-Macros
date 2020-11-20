'Used to sort specific data into the correct columns'

Sub Conformance_MetricsSO_Stats()

    Application.ScreenUpdating = False

    Columns("L:L").Cut
    Columns("B:B").Insert Shift:=xlToRight
    Columns("M:O").Cut
    Columns("C:C").Insert Shift:=xlToRight
    Columns("Q:R").Cut
    Columns("F:F").Insert Shift:=xlToRight
    Columns("H:AJ").Delete Shift:=xlToLeft
    Cells.Select
    Cells.EntireColumn.AutoFit
    Range("A1").Select

    Application.ScreenUpdating = True

End Sub
