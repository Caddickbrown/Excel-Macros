Sub MIXUP()

  Application.Calculation = xlCalculationManual
  Application.ScreenUpdating = False
  Application.DisplayStatusBar = False
  Application.EnableEvents = False

  Rows("1:1").AutoFilter
  ActiveWorkbook.Worksheets("ShopOrderOperations 191112-1516").AutoFilter.Sort. _
      SortFields.Clear
  ActiveWorkbook.Worksheets("ShopOrderOperations 191112-1516").AutoFilter.Sort. _
      SortFields.Add2 Key:=Range("D1"), SortOn:=xlSortOnValues, Order:= _
      xlAscending, DataOption:=xlSortNormal
  With ActiveWorkbook.Worksheets("ShopOrderOperations 191112-1516").AutoFilter. _
      Sort
      .Header = xlYes
      .MatchCase = False
      .Orientation = xlTopToBottom
      .SortMethod = xlPinYin
      .Apply
  End With
  Sheets("ShopOrderOperations 191112-1516").Select
  Sheets("ShopOrderOperations 191112-1516").Copy After:=Sheets(1)

  Application.EnableEvents = True
  Application.DisplayStatusBar = True
  Application.ScreenUpdating = True
  Application.Calculation = xlCalculationAutomatic

End Sub
