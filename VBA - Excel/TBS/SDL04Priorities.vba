'Used to sort specific data into the correct columns'

Sub SDL04_Priorities()

    Application.ScreenUpdating = False

    Columns("H:H").Cut
    Columns("C:C").Insert Shift:=xlToRight
    Columns("J:J").Cut
    Columns("D:D").Insert Shift:=xlToRight
    Columns("I:I").Cut
    Columns("E:E").Insert Shift:=xlToRight
    Columns("F:F").Select
    Range(Selection, Selection.End(xlToRight)).ClearContents
    Cells.Select
    Cells.EntireColumn.AutoFit
    Range("A1").Select

    Application.ScreenUpdating = True

End Sub
