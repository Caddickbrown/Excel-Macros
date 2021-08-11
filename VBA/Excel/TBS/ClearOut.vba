'Used to Delete out unneeded columns and to Sort data'

Sub ClearOut()

    Application.Calculation = xlCalculationManual
    Application.ScreenUpdating = False
    Application.DisplayStatusBar = False
    Application.EnableEvents = False

    Range("D:D,G:H,J:AC,AE:BM").Delete Shift:=xlToLeft
    Cells.Select
    Cells.EntireColumn.AutoFit
    Columns("G:G").Cut
    Columns("F:F").Insert Shift:=xlToRight
    Range("A2").Select

    Application.EnableEvents = True
    Application.DisplayStatusBar = True
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic

End Sub
