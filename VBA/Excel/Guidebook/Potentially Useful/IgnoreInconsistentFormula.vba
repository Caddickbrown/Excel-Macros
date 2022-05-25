Sub IGNOREINCONSISTENT()
  Dim r As Range: Set r = Range("A:AAA")
  Dim cel As Range

  For Each cel In r
    cel.Errors(9).Ignore = True
  Next cel

End Sub
