

Sub ShopOrderStatus_Reagent_Supply()
'
' ShopOrderStatus_Reagent_Supply Macro
'

'
    Columns("X:X").Select
    Selection.Cut
    Columns("D:D").Select
    Selection.Insert Shift:=xlToRight
    Columns("H:H").Select
    Selection.Cut
    Columns("F:F").Select
    Selection.Insert Shift:=xlToRight
    Columns("K:L").Select
    Selection.Cut
    Columns("G:G").Select
    Selection.Insert Shift:=xlToRight
    Columns("H:H").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Columns("J:CI").Select
    Selection.Delete Shift:=xlToLeft
    Columns("A:J").Select
    Columns("A:J").EntireColumn.AutoFit
    Columns("D:D").Select
    Selection.Replace What:="SDLD4 - Reagent Supply - Cal/Cntl", Replacement:= _
        "SDLD4", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, _
        SearchFormat:=False, ReplaceFormat:=False
    Selection.Replace What:="SDL04 - Reagent Supply", Replacement:="SDL04", _
        LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:= _
        False, ReplaceFormat:=False
    Selection.Replace What:="SDLC4 - Antisera Optimisation", Replacement:= _
        "SDLC4", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, _
        SearchFormat:=False, ReplaceFormat:=False
    Cells.Select
    Cells.EntireColumn.AutoFit
    Range("C10").Select
End Sub
