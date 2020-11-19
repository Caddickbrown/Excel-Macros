'This code is used to prep the Door Planning Sheet for data entry. '

Sub Door_Sheet_Prep()
'
' Door_Sheet_Prep Macro
'

'

    Application.Calculation = xlCalculationManual
    Application.ScreenUpdating = False
    Application.DisplayStatusBar = False
    Application.EnableEvents = False

    Sheets("PREMDOR DATA DUMP").Select
    Range("C3:O2000").Select
    Selection.ClearContents
    Range("A1").Select
    Sheets("JELDWEN DATA DUMP").Select
    Range("B2000:T2000").Select
    Selection.Copy
    Range("B1:T2000").Select
    ActiveSheet.Paste
    Range("A1").Select
    Sheets("FCAST SALES DUMP").Select
    Range("C2").Select
    Range("C2:AL2000").Select
    Selection.ClearContents
    Range("A1").Select
    Sheets("LOOK UPS").Select
    Range("K1").Copy
    Range("A1").Select
    Sheets("TRACKER").Select
    Range("BH1").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range("R2").Select
    Selection.Copy
    Range("Q2").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range("M2").Select
    Range("M2:M77").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("L2").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range("A1").Select

    Application.EnableEvents = True
    Application.DisplayStatusBar = True
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic

End Sub
