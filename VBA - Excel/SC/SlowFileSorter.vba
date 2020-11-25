

Sub MonthlyStockOnHandProcessor()

    Dim LastRow As Long

    Application.EnableEvents = False
    Application.DisplayStatusBar = False
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    Sheets("3 - KREP004P3").Select
    Range("AH1").FormulaR1C1 = "Macro"
    LastRow = Cells(Rows.Count, 1).End(xlUp).Row
    Range("AH2:AH" & LastRow).Formula = "=IF(SUM(RC[-25],RC[-19]:RC[-11])>0,""Keep"",""Kill"")"

    Columns("A:AH").AutoFilter
    ActiveSheet.Range("$A$1:$AH$999999").AutoFilter Field:=34, Criteria1:="Kill"

    Rows("2:999999").Select

    Selection.Delete Shift:=xlUp
    ActiveSheet.ShowAllData

    Sheets("Summary").Select
    ActiveSheet.PivotTables("PivotTable1").PivotCache.Refresh
    Sheets("3 - KREP004P3").Select

    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    Application.DisplayStatusBar = True
    Application.EnableEvents = True

End Sub
