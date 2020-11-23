'This VBA code will create a function to get the numeric part from a string

Function GetNumeric(CellRef As String)

  Dim StringLength As Integer
  StringLength = Len(CellRef)
  For i = 1 To StringLength
  If IsNumeric(Mid(CellRef, i, 1)) Then Result = Result & Mid(CellRef, i, 1)
  Next i
  GetNumeric = Result

End Function
