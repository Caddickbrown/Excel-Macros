'This code will save the entire workbook as PDF
Sub SaveWorkshetAsPDF()
ThisWorkbook.ExportAsFixedFormat xlTypePDF, "C:UsersSumitDesktopTest" & ThisWorkbook.Name & ".pdf"
End Sub

'You will have to change the folder location to use this code.'
