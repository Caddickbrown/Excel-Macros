'This was a macro for sorting Calendar data that had been exported from Outlook'

Sub CALSORT2()
'
' CALSORT2 Macro
'
'
    Application.ScreenUpdating = False
    Columns("B:B").Select
    Selection.Cut
    Columns("A:A").Select
    Selection.Insert Shift:=xlToRight
    Columns("D:D").Select
    Selection.Delete Shift:=xlToLeft
    Columns("E:E").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.ClearContents
    Range("D1").Select
    Selection.End(xlToLeft).Select
    Cells.Select
    Cells.EntireColumn.AutoFit
    Range("B1").Select
    Selection.AutoFilter
    ActiveSheet.Range("$A$1:$D$873").AutoFilter Field:=2, Criteria1:=Array( _
        "LUNCH", "LUNCH ONE", "LUNCH TWO"), Operator:=xlFilterValues
    Rows("2:2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlUp)).Select
    Rows("2:2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Delete Shift:=xlUp
    ActiveSheet.ShowAllData
    Range("A1").Select
    Application.ScreenUpdating = True
End Sub
