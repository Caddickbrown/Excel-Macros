

Sub ClearOut()
'
' ClearOut Macro
'
'
    Application.ScreenUpdating = False
    Range("D:D,G:H,J:AC,AE:BM").Select
    Range("AE1").Activate
    Selection.Delete Shift:=xlToLeft
    ActiveWindow.ScrollColumn = 16
    ActiveWindow.ScrollColumn = 15
    ActiveWindow.ScrollColumn = 14
    ActiveWindow.ScrollColumn = 13
    ActiveWindow.ScrollColumn = 12
    ActiveWindow.ScrollColumn = 11
    ActiveWindow.ScrollColumn = 10
    ActiveWindow.ScrollColumn = 9
    ActiveWindow.ScrollColumn = 8
    ActiveWindow.ScrollColumn = 7
    ActiveWindow.ScrollColumn = 6
    ActiveWindow.ScrollColumn = 5
    ActiveWindow.ScrollColumn = 4
    ActiveWindow.ScrollColumn = 3
    ActiveWindow.ScrollColumn = 2
    ActiveWindow.ScrollColumn = 1
    Cells.Select
    Cells.EntireColumn.AutoFit
    Columns("G:G").Select
    Selection.Cut
    Columns("F:F").Select
    Selection.Insert Shift:=xlToRight
    Range("A2").Select
    Application.ScreenUpdating = True
End Sub
