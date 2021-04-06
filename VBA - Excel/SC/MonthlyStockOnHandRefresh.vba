'Used to prep the Monthly Stock on Hand Sheet for data entry.''

Sub MonthlyStockOnHandRefresh()

    Application.Calculation = xlCalculationManual
    Application.ScreenUpdating = False
    Application.DisplayStatusBar = False
    Application.EnableEvents = False

    Sheets("1 - OUTSTPO").Select
    Range("A2:N2").Select
    Range(Selection, Selection.End(xlDown)).ClearContents
    Range("A2").Select
    Sheets("2 - KREP005DV1").Select
    Range("A2:S2").Select
    Range(Selection, Selection.End(xlDown)).ClearContents
    Range("A2").Select
    Sheets("3 - KREP004P3").Select
    Range("A2:J2").Select
    Range(Selection, Selection.End(xlDown)).ClearContents
    Range("M1").Copy
    Range("L1").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range("A2").Select
    Sheets("Summary").Select
    Range("A1").Select

    Application.EnableEvents = True
    Application.DisplayStatusBar = True
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic

End Sub
