

Sub FileSorter()

  Application.Calculation = xlCalculationManual
  Application.ScreenUpdating = False
  Application.DisplayStatusBar = False
  Application.EnableEvents = False

  Dim LastRow As Long

  Range("AH1").Select
  ActiveCell.FormulaR1C1 = "Macro"
  Range("AH2").Select
  ActiveCell.FormulaR1C1 = "=IF(SUM(RC[-25],RC[-20]:RC[-12])>0,""Keep"",""Kill"")"
  Selection.AutoFill Destination:=Range("AI2:AI31350")
  LastRow = Cells(Rows.Count, 1).End(xlUp).Row
  Range("B2:B" & LastRow).Formula = "=RC1+1"
