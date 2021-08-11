'This was a macro for sorting Calendar data that had been exported from Outlook'

Sub CALSORT2()

    Application.ScreenUpdating = False

    Columns("B:B").Cut
    Columns("A:A").Insert Shift:=xlToRight
    Columns("D:D").Delete Shift:=xlToLeft
    Columns("E:E").Select
    Range(Selection, Selection.End(xlToRight)).ClearContents
    Range("D1").End(xlToLeft).EntireColumn.AutoFit
    Range("B1").AutoFilter
    ActiveSheet.Range("$A$1:$D$873").AutoFilter Field:=2, Criteria1:=Array( _
        "LUNCH", "LUNCH ONE", "LUNCH TWO"), Operator:=xlFilterValues
    Rows("2:2").Select
    Range(Selection, Selection.End(xlDown)).Delete Shift:=xlUp
    ActiveSheet.ShowAllData
    Range("A1").Select

    Application.ScreenUpdating = True

End Sub
