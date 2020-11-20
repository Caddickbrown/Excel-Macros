'Refreshes CTP Data - not entirely sure how'

Sub CTP_REAPPLY_BUTTON()

    ActiveSheet.AutoFilter.ApplyFilter
    With ActiveSheet.AutoFilter.Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
End Sub
