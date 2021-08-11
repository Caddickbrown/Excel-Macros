'This code will unhide all sheets in the workbook'

Sub Unhide_All_Sheets()

  Application.Calculation = xlCalculationManual
  Application.ScreenUpdating = False
  Application.DisplayStatusBar = False
  Application.EnableEvents = False

  Dim wks As Worksheet
  For Each wks In ActiveWorkbook.Worksheets
      wks.Visible = xlSheetVisible
  Next wks

  Application.EnableEvents = True
  Application.DisplayStatusBar = True
  Application.ScreenUpdating = True
  Application.Calculation = xlCalculationAutomatic

End Sub
