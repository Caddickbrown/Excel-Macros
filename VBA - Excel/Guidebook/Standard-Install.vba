'This is a group of Macros that I install as standard to give extra functionality in Excel. They can be found at other points in this guide, but it's useful to have a single place where I can just cut and paste them in'
'Readability isn't as important here and there is no need for explainations as such - hence these will be compacted and not explained as they will be elsewhere.
Sub Unhide_All_Sheets()
    Application.ScreenUpdating = False
    Dim wks As Worksheet
    For Each wks In ActiveWorkbook.Worksheets
        wks.Visible = xlSheetVisible
    Next wks
    Application.ScreenUpdating = True
End Sub
