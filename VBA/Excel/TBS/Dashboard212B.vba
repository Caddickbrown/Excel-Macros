'Sorts data by date? Not sure if I have the info - need to see if I can remember details of how it worked'

Sub Dashboard212B()

  Application.Calculation = xlCalculationManual
  Application.ScreenUpdating = False
  Application.DisplayStatusBar = False
  Application.EnableEvents = False

  Rows("1:1").AutoFilter
  ActiveWorkbook.Worksheets(1).AutoFilter.Sort. _
      SortFields.Clear
  ActiveWorkbook.Worksheets(1).AutoFilter.Sort. _
      SortFields.Add2 Key:=Range("D1"), SortOn:=xlSortOnValues, Order:= _
      xlAscending, DataOption:=xlSortNormal
  With ActiveWorkbook.Worksheets(1).AutoFilter. _
      Sort
      .Header = xlYes
      .MatchCase = False
      .Orientation = xlTopToBottom
      .SortMethod = xlPinYin
      .Apply
  End With
  ActiveWorkbook.Worksheets(1).AutoFilter.Sort. _
      SortFields.Clear
  ActiveWorkbook.Worksheets(1).AutoFilter.Sort. _
      SortFields.Add2 Key:=Range("A1"), SortOn:=xlSortOnValues, Order:= _
      xlAscending, DataOption:=xlSortNormal
  With ActiveWorkbook.Worksheets(1).AutoFilter. _
      Sort
      .Header = xlYes
      .MatchCase = False
      .Orientation = xlTopToBottom
      .SortMethod = xlPinYin
      .Apply
  End With
  Columns("G:G").Cut
  Columns("B:B").Insert Shift:=xlToRight
  Columns("H:H").Cut
  Columns("C:C").Insert Shift:=xlToRight
  Columns("K:K").Cut
  Columns("D:D").Insert Shift:=xlToRight
  Columns("BI:BI").Cut
  Columns("E:E").Insert Shift:=xlToRight
  Columns("CF:CF").Cut
  Columns("F:F").Insert Shift:=xlToRight
  Columns("N:N").Cut
  Columns("F:F").Insert Shift:=xlToRight
  Sheets(1).Copy After:=Sheets(1)
  Sheets(1).Select
  Columns("H:H").Select
  Range(Selection, Selection.End(xlToRight)).ClearContents
  Cells.Select
  Cells.EntireColumn.AutoFit
  Sheets(2).Select
  Rows("1:1").Select
  ActiveWorkbook.Worksheets(2).AutoFilter.Sort. _
      SortFields.Clear
  ActiveWorkbook.Worksheets(2).AutoFilter.Sort. _
      SortFields.Add2 Key:=Range("J1"), SortOn:=xlSortOnValues, Order:= _
      xlDescending, DataOption:=xlSortNormal
  With ActiveWorkbook.Worksheets(2).AutoFilter. _
      Sort
      .Header = xlYes
      .MatchCase = False
      .Orientation = xlTopToBottom
      .SortMethod = xlPinYin
      .Apply
  End With
  ActiveWorkbook.Worksheets(2).AutoFilter.Sort. _
      SortFields.Clear
  ActiveWorkbook.Worksheets(2).AutoFilter.Sort. _
      SortFields.Add2 Key:=Range("A1"), SortOn:=xlSortOnValues, Order:= _
      xlAscending, DataOption:=xlSortNormal
  With ActiveWorkbook.Worksheets(2).AutoFilter. _
      Sort
      .Header = xlYes
      .MatchCase = False
      .Orientation = xlTopToBottom
      .SortMethod = xlPinYin
      .Apply
  End With
  Columns("H:H").Select
  Range(Selection, Selection.End(xlToRight)).ClearContents
  Cells.Select
  Cells.EntireColumn.AutoFit
  Worksheets(2).Range("G:G").Copy Worksheets(1).Range("G:G")
  Sheets(1).Select
  ActiveSheet.Range("$A:$CI").RemoveDuplicates Columns:=1, Header:= _
      xlYes

  Application.EnableEvents = True
  Application.DisplayStatusBar = True
  Application.ScreenUpdating = True
  Application.Calculation = xlCalculationAutomatic

End Sub
