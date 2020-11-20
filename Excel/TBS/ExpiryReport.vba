

Sub Expiry_Report()
'
' Expiry_Report Macro
'
'
    Application.ScreenUpdating = False
    Columns("D:D").Select
    Selection.Cut
    Columns("B:B").Select
    Selection.Insert Shift:=xlToRight
    Columns("H:H").Select
    Selection.Cut
    Columns("F:F").Select
    Selection.Insert Shift:=xlToRight
    Cells.Select
    Cells.EntireColumn.AutoFit
    Selection.ColumnWidth = 50.29
    Cells.EntireColumn.AutoFit
    Range("B13").Select
    Columns("B:B").ColumnWidth = 49.86
    Columns("J:J").Select
    Selection.Cut
    Columns("H:H").Select
    Selection.Insert Shift:=xlToRight
    Columns("J:J").Select
    Selection.Cut
    Columns("I:I").Select
    Selection.Insert Shift:=xlToRight
    Columns("J:J").Select
    Columns("J:BQ").Select
    Selection.Delete Shift:=xlToLeft
    Range("C1").Select
    Range("A1").Select
    Application.ScreenUpdating = True
End Sub
