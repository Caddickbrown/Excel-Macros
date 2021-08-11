'This code is used to prep the Door Planning Sheet for data entry. '

Sub ProfilesSheetPrep()

    Application.Calculation = xlCalculationManual
    Application.ScreenUpdating = False
    Application.DisplayStatusBar = False
    Application.EnableEvents = False

    'Clear Old Branch Data'
    Sheets("Branch data dump").Select
    Range("A4").Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlToRight)).ClearContents
    Range("A1").Select
    'Sales Data - Formula Rollover'
    Sheets("Master File").Select
    Range("BF4:BF25").Copy
    Range("BG4").PasteSpecial Paste:=xlPasteFormulas, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
    'Paste Values'
    Range("BF4:BF11").Copy
    Range("BF4").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range("BF13:BF19").Copy
    Range("BF13").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range("BF21:BF23").Copy
    Range("BF21").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    'Sales Data - Change Colour'
    Range("BG3").Select
    Application.CutCopyMode = False
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 65535
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    'Paste Date'
    Range("AV2").Copy
    Range("AV2").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Columns("AU:AU").Select
    Selection.Delete Shift:=xlToLeft
    'Clear out old Stock Data'
    Range("J4:J11").ClearContents
    Range("J13:J19").ClearContents
    Range("J21:J23").ClearContents
    Range("L4:M11").ClearContents
    Range("L13:M19").ClearContents
    Range("L21:M23").ClearContents
    Range("AS4:AS11").ClearContents
    Range("AS13:AS19").ClearContents
    Range("AS21:AS23").ClearContents

    Range("A1").Select

    Application.EnableEvents = True
    Application.DisplayStatusBar = True
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic

End Sub
