'Used to sort specific data into the correct columns'

Sub Conformance_ReviewSODelayType()

  Application.Calculation = xlCalculationManual
  Application.ScreenUpdating = False
  Application.DisplayStatusBar = False
  Application.EnableEvents = False

  Columns("CB:CB").Cut
  Columns("B:B").Insert Shift:=xlToRight
  Columns("C:CI").Delete Shift:=xlToLeft
  Range("A1").Select

  Application.EnableEvents = True
  Application.DisplayStatusBar = True
  Application.ScreenUpdating = True
  Application.Calculation = xlCalculationAutomatic

End Sub
