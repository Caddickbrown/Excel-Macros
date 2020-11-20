

Sub SDLA3ShopOrderStatus()
'
' SDLA3ShopOrderStatus Macro
'
'
    Application.ScreenUpdating = False
    Range("E:F,H:I").Select
    Range("H1").Activate
    Selection.Delete Shift:=xlToLeft
    Columns("H:I").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.ClearContents
    Cells.Select
    Cells.EntireColumn.AutoFit
    Range("A1:G1").Select
    Selection.AutoFilter
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
    Application.ScreenUpdating = True
End Sub
