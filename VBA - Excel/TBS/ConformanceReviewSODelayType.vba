'Used to sort specific data into the correct columns'

Sub Conformance_ReviewSODelayType()
'
' Conformance_ReviewSODelayType Macro
'
'
    Application.ScreenUpdating = False
    Range("A1").Select
    Columns("CB:CB").Select
    Application.CutCopyMode = False
    Selection.Cut
    Columns("B:B").Select
    Selection.Insert Shift:=xlToRight
    Columns("C:C").Select
    Columns("C:CI").Select
    Selection.Delete Shift:=xlToLeft
    Range("A1").Select
    Application.ScreenUpdating = True
End Sub
