

Sub MonthlyStockOnHandProcessor()

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

    If ActiveSheet.AutoFilterMode Then
    ActiveSheet.AutoFilterMode = False
    End If

    Columns("A:AH").AutoFilter
    ActiveSheet.Range("$A$1:$AH$99999").AutoFilter Field:=34, Criteria1:="Kill"

    Range("A2:AH99999").SpecialCells(xlCellTypeVisible).EntireRow.Delete

    Range("A1").Select
    ActiveSheet.ShowAllData
    ActiveSheet.AutoFilterMode = False
    Application.Calculation = xlCalculationManual

    Sheets("Summary").Select
    ActiveSheet.PivotTables("PivotTable1").PivotCache.Refresh
    Sheets("3 - KREP004P3").Select
    Range("A1").Select

    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    Application.DisplayStatusBar = True
    Application.EnableEvents = True

End Sub
