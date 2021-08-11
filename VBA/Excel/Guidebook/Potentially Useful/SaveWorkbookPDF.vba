'This code will save the entire workbook as PDF
'You will have to change the folder location to use this code.

Sub SaveWorksheetAsPDF()

  ThisWorkbook.ExportAsFixedFormat xlTypePDF, "C:UsersSumitDesktopTest" & ThisWorkbook.Name & ".pdf"

End Sub
