'Not sure'

Sub NotesMacro()

  Application.Calculation = xlCalculationManual
  Application.ScreenUpdating = False
  Application.DisplayStatusBar = False
  Application.EnableEvents = False

  Worksheets("Gantt").Select
  Range("A1").Activate
  Worksheets("Notes").ShowDataForm

  Application.EnableEvents = True
  Application.DisplayStatusBar = True
  Application.ScreenUpdating = True
  Application.Calculation = xlCalculationAutomatic

End Sub
