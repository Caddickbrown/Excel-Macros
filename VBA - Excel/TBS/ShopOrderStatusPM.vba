'Used to sort specific data into the correct columns'

Sub SHOP_ORDER_STATUS_PM()
'
' SHOP_ORDER_STATUS_PM Macro
'

'
    Columns("G:G").Select
    Selection.Cut
    Columns("E:E").Select
    Selection.Insert Shift:=xlToRight
    Columns("J:K").Select
    Selection.Cut
    Columns("F:F").Select
    Selection.Insert Shift:=xlToRight
    Columns("G:G").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Columns("I:I").Select
    Columns("I:CH").Select
    Selection.ClearContents
    Columns("A:H").Select
    Columns("A:H").EntireColumn.AutoFit
    Range("A1").Select
End Sub
