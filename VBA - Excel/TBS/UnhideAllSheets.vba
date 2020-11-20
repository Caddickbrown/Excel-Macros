

Sub Unhide_All_Sheets()
    Application.ScreenUpdating = False
    Dim wks As Worksheet

    For Each wks In ActiveWorkbook.Worksheets
        wks.Visible = xlSheetVisible
    Next wks
    Application.ScreenUpdating = True
End Sub
