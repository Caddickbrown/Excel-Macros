

Sub Dashboard212B()
'
' Dashboard212B Macro
'
'
    Application.ScreenUpdating = False
    Rows("1:1").Select
    Selection.AutoFilter
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
    Columns("G:G").Select
    Selection.Cut
    Columns("B:B").Select
    Selection.Insert Shift:=xlToRight
    Columns("H:H").Select
    Selection.Cut
    Columns("C:C").Select
    Selection.Insert Shift:=xlToRight
    Columns("K:K").Select
    Selection.Cut
    Columns("D:D").Select
    Selection.Insert Shift:=xlToRight
    Columns("BI:BI").Select
    Selection.Cut
    Columns("E:E").Select
    Selection.Insert Shift:=xlToRight
    Columns("CF:CF").Select
    Selection.Cut
    Columns("F:F").Select
    Selection.Insert Shift:=xlToRight
    Columns("N:N").Select
    Selection.Cut
    Columns("F:F").Select
    Selection.Insert Shift:=xlToRight
    Sheets(1).Select
    Sheets(1).Copy After:=Sheets(1)
    Sheets(1).Select
    Columns("H:H").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.ClearContents
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
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.ClearContents
    Cells.Select
    Cells.EntireColumn.AutoFit
    Worksheets(2).Range("G:G").Copy Worksheets(1).Range("G:G")
    Sheets(1).Select
    ActiveSheet.Range("$A:$CI").RemoveDuplicates Columns:=1, Header:= _
        xlYes
    Application.ScreenUpdating = True
End Sub
