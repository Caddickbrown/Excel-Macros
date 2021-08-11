'This is a group of Macros that I install as standard to give extra functionality in Excel. They can be found at other points in this guide, but it's useful to have a single place where I can just cut and paste them in
'Readability isn't as important here and there is no need for explainations as such - hence these will be compacted and explained elsewhere.
Sub Unhide_All_Sheets()
    Application.ScreenUpdating = False
    Dim wks As Worksheet
    For Each wks In ActiveWorkbook.Worksheets
        wks.Visible = xlSheetVisible
    Next wks
    Application.ScreenUpdating = True
End Sub
Function SumByColour(CellColor As Range, rRange As Range)
  Dim cSum As Long
  Dim ColIndex As Integer
  ColIndex = CellColor.Interior.ColorIndex
  For Each cl In rRange
    If cl.Interior.ColorIndex = ColIndex Then
    cSum = WorksheetFunction.Sum(cl, cSum)
  End If
  Next cl
  SumByColour = cSum
End Function
Function CountColourIf(rSample As Range, rArea As Range) As Long
    Dim rAreaCell As Range
    Dim lMatchColor As Long
    Dim lCounter As Long
    lMatchColor = rSample.Interior.Color
    For Each rAreaCell In rArea
        If rAreaCell.Interior.Color = lMatchColor Then
            lCounter = lCounter + 1
        End If
    Next rAreaCell
    CountColourIf = lCounter
End Function
Function GetNumeric(CellRef As String)
  Dim StringLength As Integer
  StringLength = Len(CellRef)
  For i = 1 To StringLength
  If IsNumeric(Mid(CellRef, i, 1)) Then Result = Result & Mid(CellRef, i, 1)
  Next i
  GetNumeric = Result
End Function
Function GetText(CellRef As String)
  Dim StringLength As Integer
  StringLength = Len(CellRef)
  For i = 1 To StringLength
  If Not (IsNumeric(Mid(CellRef, i, 1))) Then Result = Result & Mid(CellRef, i, 1)
  Next i
  GetText = Result
End Function
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
