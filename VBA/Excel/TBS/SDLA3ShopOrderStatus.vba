

Sub SDLA3ShopOrderStatus()

  Application.Calculation = xlCalculationManual
  Application.ScreenUpdating = False
  Application.DisplayStatusBar = False
  Application.EnableEvents = False

    Range("E:F,H:I").Delete Shift:=xlToLeft
    Columns("H:I").Select
    Range(Selection, Selection.End(xlToRight)).ClearContents
    Cells.Select
    Cells.EntireColumn.AutoFit
    Range("A1:G1").AutoFilter
    ActiveWorkbook.Worksheets(1).AutoFilter.Sort.SortFields _
        .Clear
    ActiveWorkbook.Worksheets(1).AutoFilter.Sort.SortFields _
        .Add2 Key:=Range("F1"), SortOn:=xlSortOnValues, Order:=xlAscending, _
        DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets(1).AutoFilter.Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    Range("A1").Select

  Application.EnableEvents = True
  Application.DisplayStatusBar = True
  Application.ScreenUpdating = True
  Application.Calculation = xlCalculationAutomatic

End Sub
