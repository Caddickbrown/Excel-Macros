

Sub FileSorter()

  Dim LastRow As Long

  Application.EnableEvents = False
  Application.DisplayStatusBar = False
  Application.ScreenUpdating = False
  Application.Calculation = xlCalculationManual


  Range("AH1").FormulaR1C1 = "Macro"
  LastRow = Cells(Rows.Count, 1).End(xlUp).Row
  Range("AH2:AH" & LastRow).Formula = "=IF(SUM(RC[-25],RC[-20]:RC[-12])>0,""Keep"",""Kill"")"

  Columns("A:AH").AutoFilter
  ActiveSheet.Range("$A$1:$AH$999999").AutoFilter Field:=34, Criteria1:= "Kill"

  Rows("2:2").Select
  Range (Selection, Selection.End(x1Down)).Delete Shift:=xlUp
  ActiveSheet.ShowAllData

  Application.Calculation = xlCalculationAutomatic
  Application.ScreenUpdating = True
  Application.DisplayStatusBar = True
  Application.EnableEvents = True

End Sub
