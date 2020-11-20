'Used to sort specific data into the correct columns'

Sub CTP_DATA()

    Columns("G:G").Cut
    Columns("A:A").Insert Shift:=xlToRight
    Columns("BI:BI").Cut
    Columns("C:C").Insert Shift:=xlToRight
    Columns("CF:CF").Cut
    Columns("D:D").Insert Shift:=xlToRight
    Columns("M:M").Cut
    Columns("E:E").Insert Shift:=xlToRight
    Columns("H:H").Cut
    Columns("F:F").Insert Shift:=xlToRight
    Columns("G:CI").Delete Shift:=xlToLeft
    Cells.EntireColumn.AutoFit
    Range("A1").Select

End Sub
