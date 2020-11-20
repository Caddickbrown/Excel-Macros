'This code will Save the File With a Timestamp in its name
Sub SaveWorkbookWithTimeStamp()
Dim timestamp As String
timestamp = Format(Date, "dd-mm-yyyy") & "_" & Format(Time, "hh-ss")
ThisWorkbook.SaveAs "C:UsersUsernameDesktopWorkbookName" & timestamp
End Sub
