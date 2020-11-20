

Sub CTP_Data_FullReDeux()

    Application.ScreenUpdating = False

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
    Columns("G:G").Select
    Range(Selection, Selection.End(xlToRight)).ClearContents
    Cells.EntireColumn.AutoFit
    Sheets(1).Copy After:=Sheets(1)
    Range("A1:F1").Select
    Range("F1").Activate
    Selection.AutoFilter
    ActiveWorkbook.Worksheets(2).AutoFilter.Sort. _
        SortFields.Clear
    ActiveWorkbook.Worksheets(2).AutoFilter.Sort. _
        SortFields.Add2 Key:=Range("F1"), SortOn:=xlSortOnValues, Order:= _
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
        SortFields.Add2 Key:=Range("B1"), SortOn:=xlSortOnValues, Order:= _
        xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets(2).AutoFilter. _
        Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    ActiveSheet.Range("$A:$F").RemoveDuplicates Columns:=2, Header:=xlYes
    Sheets(1).Select
    Range("A1:F1").Select
    Range("F1").Activate
    Selection.AutoFilter
    ActiveWorkbook.Worksheets(1).AutoFilter.Sort. _
        SortFields.Clear
    ActiveWorkbook.Worksheets(1).AutoFilter.Sort. _
        SortFields.Add2 Key:=Range("F1"), SortOn:=xlSortOnValues, Order:= _
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
        SortFields.Add2 Key:=Range("B1"), SortOn:=xlSortOnValues, Order:= _
        xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets(1).AutoFilter. _
        Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    ActiveSheet.Range("$A:$F").RemoveDuplicates Columns:=2, Header:=xlYes
    Sheets(2).Select
    Columns("D:D").Select
    Selection.Copy
    Sheets(1).Select
    Columns("D:D").Select
    ActiveSheet.Paste
    Range("A1").Select
    Sheets(2).Select
    Application.SendKeys "{ENTER}"
    ActiveWindow.SelectedSheets.Delete
    Columns("A:F").Select
    Selection.Copy
    Range("A1").Select

    Application.ScreenUpdating = True

End Sub
