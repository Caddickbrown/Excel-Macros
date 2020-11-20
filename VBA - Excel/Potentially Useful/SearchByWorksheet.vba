'VBA: Search by worksheet name'

Sub SearchSheetName()

  Dim xName As String
  Dim xFound As Boolean

  xName = InputBox("Enter sheet name to find in workbook:", "Sheet search")
  If xName = "" Then Exit Sub
  On Error Resume Next
  ActiveWorkbook.Sheets(xName).Select
  xFound = (Err = 0)
  On Error GoTo 0
  If xFound Then
    MsgBox "Sheet '" & xName & "' has been found and selected!"
    Else
    MsgBox "The sheet '" & xName & "' could not be found in this workbook!"
  End If

End Sub
