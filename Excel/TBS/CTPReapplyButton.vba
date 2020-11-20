

Sub CTP_REAPPLY_BUTTON()
'
' reapply Macro
'

'
    ActiveSheet.AutoFilter.ApplyFilter
    With ActiveSheet.AutoFilter.Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
End Sub
