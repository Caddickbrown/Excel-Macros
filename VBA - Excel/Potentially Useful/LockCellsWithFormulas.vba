'This macro code will lock all the cells with formulas
Sub LockCellsWithFormulas()
With ActiveSheet
   .Unprotect
   .Cells.Locked = False
   .Cells.SpecialCells(xlCellTypeFormulas).Locked = True
   .Protect AllowDeletingRows:=True
End With
End Sub
