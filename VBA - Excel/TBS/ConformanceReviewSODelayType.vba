'Used to sort specific data into the correct columns'

Sub Conformance_ReviewSODelayType()

    Application.ScreenUpdating = False

    Columns("CB:CB").Cut
    Columns("B:B").Insert Shift:=xlToRight
    Columns("C:CI").Delete Shift:=xlToLeft
    Range("A1").Select

    Application.ScreenUpdating = True

End Sub
