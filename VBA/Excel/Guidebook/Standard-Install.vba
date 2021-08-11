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
Sub UnhideAllRowsColumns()
  Columns.EntireColumn.Hidden = False
  Rows.EntireRow.Hidden = False
End Sub
Sub MergeExcelFiles()
    Application.Calculation = xlCalculationManual
    Application.ScreenUpdating = False
    Application.DisplayStatusBar = False
    Application.EnableEvents = False
    Dim fnameList, fnameCurFile As Variant
    Dim countFiles, countSheets As Integer
    Dim wksCurSheet As Worksheet
    Dim wbkCurBook, wbkSrcBook As Workbook
    fnameList = Application.GetOpenFilename(FileFilter:="Microsoft Excel Workbooks (*.xls;*.xlsx;*.xlsm),*.xls;*.xlsx;*.xlsm", Title:="Choose Excel files to merge", MultiSelect:=True)
    If (vbBoolean <> VarType(fnameList)) Then
        If (UBound(fnameList) > 0) Then
            countFiles = 0
            countSheets = 0
            Set wbkCurBook = ActiveWorkbook
            For Each fnameCurFile In fnameList
                countFiles = countFiles + 1
                Set wbkSrcBook = Workbooks.Open(Filename:=fnameCurFile)
                For Each wksCurSheet In wbkSrcBook.Sheets
                    countSheets = countSheets + 1
                    wksCurSheet.Copy after:=wbkCurBook.Sheets(wbkCurBook.Sheets.Count)
                Next
                wbkSrcBook.Close SaveChanges:=False
            Next
            MsgBox "Processed " & countFiles & " files" & vbCrLf & "Merged " & countSheets & " worksheets", Title:="Merge Excel Files"
        End If
    Else
        MsgBox "No files selected", Title:="Merge Excel Files"
    End If
    Application.EnableEvents = True
    Application.DisplayStatusBar = True
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
End Sub
