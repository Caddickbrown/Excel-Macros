'This Macro creates a "SumByColor" funsction that will count the number of cells with a defined colour

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
