'Used to sort specific data into the correct columns'

Sub SHOP_ORDER_STATUS_PM()

    Columns("G:G").Cut
    Columns("E:E").Insert Shift:=xlToRight
    Columns("J:K").Cut
    Columns("F:F").Insert Shift:=xlToRight
    Columns("G:G").Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Columns("I:CH").Select
    Selection.ClearContents
    Cells.Select
    Cells.EntireColumn.AutoFit
    Range("A1").Select

End Sub
