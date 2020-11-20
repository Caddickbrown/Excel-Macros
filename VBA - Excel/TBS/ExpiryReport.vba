'Not sure what this does - just sorts data it would seem'

Sub Expiry_Report()

    Application.ScreenUpdating = False

    Columns("D:D").Cut
    Columns("B:B").Insert Shift:=xlToRight
    Columns("H:H").Cut
    Columns("F:F").Insert Shift:=xlToRight
    Cells.Select
    Cells.EntireColumn.AutoFit
    Columns("J:J").Cut
    Columns("H:H").Insert Shift:=xlToRight
    Columns("J:J").Cut
    Columns("I:I").Insert Shift:=xlToRight
    Columns("J:BQ").Delete Shift:=xlToLeft
    Range("A1").Select

    Application.ScreenUpdating = True

End Sub
