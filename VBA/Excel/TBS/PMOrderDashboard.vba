'Moves columns around and deletes out unneeded stuff'

Sub PM_OrderDashboard()

  Application.Calculation = xlCalculationManual
  Application.ScreenUpdating = False
  Application.DisplayStatusBar = False
  Application.EnableEvents = False

  Columns("G:G").Cut
  Columns("E:E").Insert Shift:=xlToRight
  Columns("I:I").Cut
  Columns("F:F").Insert Shift:=xlToRight
  Columns("J:J").Cut
  Columns("G:G").Insert Shift:=xlToRight
  Columns("J:J").Cut
  Columns("H:H").Insert Shift:=xlToRight
  Columns("I:I").Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
  Columns("L:L").Cut
  Columns("J:J").Insert Shift:=xlToRight
  Columns("K:K").Select
  Range(Selection, Selection.End(xlToRight)).Select
  Selection.ClearContents
  Selection.End(xlToLeft).Select
  Selection.End(xlToLeft).Select
  Selection.End(xlToLeft).Select
  Cells.Select
  Cells.EntireColumn.AutoFit
  Range("A1").Select

  Application.EnableEvents = True
  Application.DisplayStatusBar = True
  Application.ScreenUpdating = True
  Application.Calculation = xlCalculationAutomatic

End Sub
