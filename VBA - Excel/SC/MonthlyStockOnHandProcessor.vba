

Sub MonthlyStockOnHandProcessor()

    Dim lastRow1 As Long

    Application.EnableEvents = False
    Application.DisplayStatusBar = False
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    Sheets("3 - KREP004P3").Select
    Range("AH1").FormulaR1C1 = "Macro"

    Range("AH2").FormulaR1C1 = "=IF(SUM(RC[-25],RC[-19]:RC[-11])>0,""Keep"",""Kill"")"
    Range("M2:AH2").Copy
    Range("M2:M99999").Select
    ActiveSheet.Paste
    Application.Calculation = xlCalculationAutomatic

'    LastRow1 = Cells(Rows.Count, 1).End(xlUp).Row
'    Range("AH2:AH" & LastRow1).Formula = "=IF(SUM(RC[-25],RC[-19]:RC[-11])>0,""Keep"",""Kill"")"

    If ActiveSheet.AutoFilterMode Then
    ActiveSheet.AutoFilterMode = False
    End If

    Columns("A:AH").AutoFilter

    Range("AH2:AH99999").Select
    Selection.Copy
    Range("AH2").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False

    Application.Calculation = xlCalculationManual

    ActiveWorkbook.ActiveSheet.AutoFilter.Sort. _
        SortFields.Clear
    ActiveWorkbook.ActiveSheet.AutoFilter.Sort. _
        SortFields.Add2 Key:=Range("AH1"), SortOn:=xlSortOnValues, Order:= _
        xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.ActiveSheet.AutoFilter. _
        Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With

    ActiveSheet.Range("$A$1:$AH$99999").AutoFilter Field:=34, Criteria1:="Kill"

    Range("A2:AH99999").SpecialCells(xlCellTypeVisible).EntireRow.Delete

    Range("A1").Select
    ActiveSheet.ShowAllData
    ActiveSheet.AutoFilterMode = False

    Sheets("Summary").Select
    ActiveSheet.PivotTables("PivotTable1").PivotCache.Refresh
    Range("A1").Select

    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    Application.DisplayStatusBar = True
    Application.EnableEvents = True

End Sub
