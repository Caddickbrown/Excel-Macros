'Surround all macros with these two blocks - it will speed up the process'
'It might not be by much, but it mounts up over time.'

Application.Calculation = xlCalculationManual
Application.ScreenUpdating = False
Application.DisplayStatusBar = False
Application.EnableEvents = False

Application.EnableEvents = True
Application.DisplayStatusBar = True
Application.ScreenUpdating = True
Application.Calculation = xlCalculationAutomatic
