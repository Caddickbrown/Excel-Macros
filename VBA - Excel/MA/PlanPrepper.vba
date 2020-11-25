'This macro spat out multiple tabs, sorted data from a data dump tab and copied into the relevant tabs'

Sub PressPlans()

    Application.Calculation = xlCalculationManual
    Application.ScreenUpdating = False
    Application.DisplayStatusBar = False
    Application.EnableEvents = False

    Sheets.Add After:=ActiveSheet
    Sheets("Sheet1").Select
    Sheets("Sheet1").Name = "12000T"
    Sheets.Add After:=ActiveSheet
    Sheets("Sheet2").Select
    Sheets("Sheet2").Name = "750T"
    Sheets.Add After:=ActiveSheet
    Sheets("Sheet3").Select
    Sheets("Sheet3").Name = "1250T"
    Sheets.Add After:=ActiveSheet
    Sheets("Sheet4").Select
    Sheets("Sheet4").Name = "25002000T"
    Sheets.Add After:=ActiveSheet
    Sheets("Sheet5").Select
    Sheets("Sheet5").Name = "30001000RR"
    Sheets.Add After:=ActiveSheet
    Sheets("Sheet6").Select
    Sheets("Sheet6").Name = "DDP"
    Sheets.Add After:=ActiveSheet
    Sheets("Sheet7").Select
    Sheets("Sheet7").Name = "LightCell"
    Sheets.Add After:=ActiveSheet
    Sheets("Sheet8").Select
    Sheets("Sheet8").Name = "Open"
    Sheets("Part Information").Select
    Columns("D:D").Select
    Selection.Cut
    Columns("A:A").Select
    Selection.Insert Shift:=xlToRight
    Columns("F:F").Select
    Selection.Cut
    Columns("B:B").Select
    Selection.Insert Shift:=xlToRight
    Columns("H:H").Select
    Selection.Cut
    Columns("C:C").Select
    Selection.Insert Shift:=xlToRight
    Columns("I:I").Select
    Selection.Cut
    Columns("D:D").Select
    Selection.Insert Shift:=xlToRight
    Columns("J:J").Select
    Selection.Cut
    Columns("E:E").Select
    Selection.Insert Shift:=xlToRight
    Columns("M:M").Select
    Selection.Cut
    Columns("F:F").Select
    Selection.Insert Shift:=xlToRight
    Columns("I:I").Select
    Selection.Cut
    Columns("G:G").Select
    Selection.Insert Shift:=xlToRight
    Columns("O:O").Select
    Selection.Cut
    Columns("H:H").Select
    Selection.Insert Shift:=xlToRight
    Columns("P:P").Select
    Selection.Cut
    Columns("I:I").Select
    Selection.Insert Shift:=xlToRight
    Columns("U:U").Select
    Selection.Cut
    Columns("J:J").Select
    Selection.Insert Shift:=xlToRight
    Columns("Q:Q").Select
    Selection.Cut
    Columns("K:K").Select
    Selection.Insert Shift:=xlToRight
    Columns("R:R").Select
    Selection.Cut
    Columns("L:L").Select
    Selection.Insert Shift:=xlToRight
    Columns("W:X").Select
    Selection.Cut
    Columns("M:M").Select
    Selection.Insert Shift:=xlToRight
    Columns("AB:AB").Select
    Selection.Cut
    Columns("O:O").Select
    Selection.Insert Shift:=xlToRight
    Columns("A:O").Select
    Selection.AutoFilter
    ActiveSheet.Range("$A$1:$O$3052").AutoFilter Field:=7, Criteria1:= _
        "12000T PRESS"
    Selection.Copy
    Sheets("12000T").Select
    Columns("A:A").Select
    ActiveSheet.Paste
    Sheets("Part Information").Select
    ActiveSheet.ShowAllData
    Columns("Z:AA").Select
    Selection.Cut
    Columns("L:L").Select
    Selection.Insert Shift:=xlToRight
    Columns("G:G").Select
    Selection.Cut
    Columns("F:F").Select
    Selection.Insert Shift:=xlToRight
    Columns("R:R").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.ClearContents
    ActiveSheet.Range("$A$1:$Q$3052").AutoFilter Field:=6, Criteria1:= _
        "750T PRESS"
    Columns("A:Q").Select
    Selection.Copy
    Sheets("750T").Select
    ActiveSheet.Paste
    Sheets("Part Information").Select
    ActiveSheet.Range("$A$1:$Q$3052").AutoFilter Field:=6, Criteria1:= _
        "1250T PRESS"
    Columns("A:Q").Select
    Application.CutCopyMode = False
    Selection.Copy
    Sheets("1250T").Select
    ActiveSheet.Paste
    Sheets("Part Information").Select
    ActiveSheet.Range("$A$1:$Q$3052").AutoFilter Field:=6, Criteria1:= _
        "=2000T PRESS", Operator:=xlOr, Criteria2:="=2500T PRESS"
    Columns("A:Q").Select
    Application.CutCopyMode = False
    Selection.Copy
    Sheets("25002000T").Select
    Range("A1").Select
    ActiveSheet.Paste
    Sheets("Part Information").Select
    Range("F11").Select
    ActiveSheet.Range("$A$1:$Q$3052").AutoFilter Field:=6, Criteria1:=Array( _
        "3000T PRESS", "HDA 1000T PRESS", "RR 80 TON RING ROLLER"), Operator:= _
        xlFilterValues
    Columns("A:Q").Select
    Application.CutCopyMode = False
    Selection.Copy
    Sheets("30001000RR").Select
    ActiveSheet.Paste
    Sheets("Part Information").Select
    ActiveSheet.Range("$A$1:$Q$3052").AutoFilter Field:=6, Criteria1:= _
        "DDP 2000 T"
    Columns("A:Q").Select
    Application.CutCopyMode = False
    Selection.Copy
    Sheets("DDP").Select
    ActiveSheet.Paste
    Sheets("Part Information").Select
    ActiveSheet.Range("$A$1:$Q$3052").AutoFilter Field:=6, Criteria1:=Array( _
        "1500T PRESS", "200T PRESS", "500T PRESS", "800T PRESS"), Operator:= _
        xlFilterValues
    Columns("A:Q").Select
    Application.CutCopyMode = False
    Selection.Copy
    Sheets("LightCell").Select
    ActiveSheet.Paste
    Sheets("Part Information").Select
    ActiveSheet.Range("$A$1:$Q$3052").AutoFilter Field:=6, Criteria1:= _
        "HDA OPEN FORGE"
    Columns("A:Q").Select
    Application.CutCopyMode = False
    Selection.Copy
    Sheets("Open").Select
    ActiveSheet.Paste
    Columns("A:Q").Select
    Selection.AutoFilter
    Cells.Select
    Cells.EntireColumn.AutoFit
    Range("A1").Select
    Sheets("LightCell").Select
    Columns("A:Q").Select
    Selection.AutoFilter
    Cells.Select
    Cells.EntireColumn.AutoFit
    Range("A1").Select
    Sheets("DDP").Select
    Columns("A:Q").Select
    Selection.AutoFilter
    Cells.Select
    Cells.EntireColumn.AutoFit
    Range("A1").Select
    Sheets("30001000RR").Select
    Columns("A:Q").Select
    Selection.AutoFilter
    Cells.Select
    Cells.EntireColumn.AutoFit
    Range("A1").Select
    Sheets("25002000T").Select
    Columns("A:Q").Select
    Selection.AutoFilter
    Cells.Select
    Cells.EntireColumn.AutoFit
    Range("A1").Select
    Sheets("1250T").Select
    Columns("A:Q").Select
    Selection.AutoFilter
    Cells.Select
    Cells.EntireColumn.AutoFit
    Range("A1").Select
    Sheets("750T").Select
    Columns("A:Q").Select
    Selection.AutoFilter
    Cells.Select
    Cells.EntireColumn.AutoFit
    Range("A1").Select
    Sheets("12000T").Select
    Columns("D:D").Select
    Application.CutCopyMode = False
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Columns("G:G").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("D1").Select
    ActiveCell.FormulaR1C1 = "Temp"
    Range("G1").Select
    ActiveCell.FormulaR1C1 = "Setup"
    Range("H1").Select
    ActiveCell.FormulaR1C1 = "Temp"
    Columns("A:R").Select
    Selection.AutoFilter
    Cells.Select
    Cells.EntireColumn.AutoFit
    Range("A1").Select
    Sheets("Part Information").Select
    ActiveSheet.Range("$A$1:$Q$3052").AutoFilter Field:=6
    Cells.Select
    Cells.EntireColumn.AutoFit
    Range("A1").Select

    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    Application.DisplayStatusBar = True
    Application.EnableEvents = True

End Sub
