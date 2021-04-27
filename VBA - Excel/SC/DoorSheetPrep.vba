'This code is used to prep the Door Planning Sheet for data entry.'

Sub Door_Sheet_Prep()

    Application.Calculation = xlCalculationManual
    Application.ScreenUpdating = False
    Application.DisplayStatusBar = False
    Application.EnableEvents = False

    'Clear Old Premdor Data'
    Sheets("PREMDOR DATA DUMP").Select
    Range("C3:Q2000").ClearContents
    Range("A1").Select
    'Clear Old Jeld-Wen Data'
    Sheets("JELDWEN DATA DUMP").Select
    Range("B2000:T2000").Copy
    Range("B1:T2000").Select
    ActiveSheet.Paste
    Range("A1").Select
    'Clear Old Sales Data'
    Sheets("FCAST SALES DUMP").Select
    Range("C2").Select
    Range("C2:AL999999").ClearContents
    Range("A1").Select
    'Copy/Paste Week Change'
    Sheets("LOOK UPS").Select
    Range("K1").Copy
    Range("A1").Select
    Sheets("TRACKER").Select
    Range("BG1").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    'Move Date along'
    Range("Q2").Copy
    Range("P2").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    'Move Stock Take Data'
    Range("M2:M58").Copy
    Range("L2").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range("M60:M73").Copy
    Range("L60").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range("M3:M58", "M60:M73").ClearContents
    Range("A1").Select

    Application.EnableEvents = True
    Application.DisplayStatusBar = True
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic

End Sub
