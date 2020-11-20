'Used to sort specific data into the correct columns'

Sub CTP_DATA()

    Columns("G:G").Select
    Selection.Cut
    Columns("A:A").Select
    Selection.Insert Shift:=xlToRight
    Columns("BI:BI").Select
    Selection.Cut
    Columns("BI:BI").Select
    Columns("C:C").Select
    Selection.Insert Shift:=xlToRight
    Columns("CF:CF").Select
    Selection.Cut
    Columns("D:D").Select
    Selection.Insert Shift:=xlToRight
    Columns("M:M").Select
    Selection.Cut
    Columns("E:E").Select
    Selection.Insert Shift:=xlToRight
    Columns("H:H").Select
    Selection.Cut
    Columns("F:F").Select
    Selection.Insert Shift:=xlToRight
    Columns("G:G").Select
    Columns("G:CI").Select
    Selection.Delete Shift:=xlToLeft
    Cells.Select
    Cells.EntireColumn.AutoFit
    Range("G5").Select
End Sub
