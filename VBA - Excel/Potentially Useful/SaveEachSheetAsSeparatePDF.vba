'This code will save each worsheet as a separate PDF
'You will have to change the folder location to use this code.

Sub SaveWorksheetAsPDF()

  Dim ws As Worksheet

  For Each ws In Worksheets
  ws.ExportAsFixedFormat xlTypePDF, "C:UsersSumitDesktopTest" & ws.Name & ".pdf"
  Next ws

End Sub
