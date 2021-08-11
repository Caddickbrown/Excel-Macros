'This macro spat out multiple tabs, sorted data from a data dump tab and copied into the relevant tabs'

Sub PressPlans()

    Application.Calculation = xlCalculationManual
    Application.ScreenUpdating = False
    Application.DisplayStatusBar = False
    Application.EnableEvents = False

    Sheets.Add After:=ActiveSheet
    Sheets("Sheet1").Name = "12000T"
    Sheets.Add After:=ActiveSheet
    Sheets("Sheet2").Name = "750T"
    Sheets.Add After:=ActiveSheet
    Sheets("Sheet3").Name = "1250T"
    Sheets.Add After:=ActiveSheet
    Sheets("Sheet4").Name = "25002000T"
    Sheets.Add After:=ActiveSheet
    Sheets("Sheet5").Name = "30001000RR"
    Sheets.Add After:=ActiveSheet
    Sheets("Sheet6").Name = "DDP"
    Sheets.Add After:=ActiveSheet
    Sheets("Sheet7").Name = "LightCell"
    Sheets.Add After:=ActiveSheet
    Sheets("Sheet8").Name = "Open"
    Sheets("Part Information").Select
    Columns("D:D").Cut
    Columns("A:A").Insert Shift:=xlToRight
    Columns("F:F").Cut
    Columns("B:B").Insert Shift:=xlToRight
    Columns("H:H").Cut
    Columns("C:C").Insert Shift:=xlToRight
    Columns("I:I").Cut
    Columns("D:D").Insert Shift:=xlToRight
    Columns("J:J").Cut
    Columns("E:E").Insert Shift:=xlToRight
    Columns("M:M").Cut
    Columns("F:F").Insert Shift:=xlToRight
    Columns("I:I").Cut
    Columns("G:G").Insert Shift:=xlToRight
    Columns("O:O").Cut
    Columns("H:H").Insert Shift:=xlToRight
    Columns("P:P").Cut
    Columns("I:I").Insert Shift:=xlToRight
    Columns("U:U").Cut
    Columns("J:J").Insert Shift:=xlToRight
    Columns("Q:Q").Cut
    Columns("K:K").Insert Shift:=xlToRight
    Columns("R:R").Cut
    Columns("L:L").Insert Shift:=xlToRight
    Columns("W:X").Cut
    Columns("M:M").Insert Shift:=xlToRight
    Columns("AB:AB").Cut
    Columns("O:O").Insert Shift:=xlToRight
    Columns("A:O").AutoFilter
    ActiveSheet.Range("$A$1:$O$3052").AutoFilter Field:=7, Criteria1:= "12000T PRESS"
    Selection.Copy
    Sheets("12000T").Select
    Columns("A:A").Select
    ActiveSheet.Paste
    Sheets("Part Information").Select
    ActiveSheet.ShowAllData
    Columns("Z:AA").Cut
    Columns("L:L").Insert Shift:=xlToRight
    Columns("G:G").Cut
    Columns("F:F").Insert Shift:=xlToRight
    Columns("R:R").Select
    Range(Selection, Selection.End(xlToRight)).ClearContents
    ActiveSheet.Range("$A$1:$Q$3052").AutoFilter Field:=6, Criteria1:= _
        "750T PRESS"
    Columns("A:Q").Copy
    Sheets("750T").Paste
    Sheets("Part Information").Select
    ActiveSheet.Range("$A$1:$Q$3052").AutoFilter Field:=6, Criteria1:= _
        "1250T PRESS"
    Columns("A:Q").Copy
    Sheets("1250T").Select
    ActiveSheet.Paste
    Sheets("Part Information").Select
    ActiveSheet.Range("$A$1:$Q$3052").AutoFilter Field:=6, Criteria1:= _
        "=2000T PRESS", Operator:=xlOr, Criteria2:="=2500T PRESS"
    Columns("A:Q").Copy
    Sheets("25002000T").Select
    Range("A1").Select
    ActiveSheet.Paste
    Sheets("Part Information").Select
    Range("F11").Select
    ActiveSheet.Range("$A$1:$Q$3052").AutoFilter Field:=6, Criteria1:=Array( _
        "3000T PRESS", "HDA 1000T PRESS", "RR 80 TON RING ROLLER"), Operator:= _
        xlFilterValues
    Columns("A:Q").Copy
    Sheets("30001000RR").Select
    ActiveSheet.Paste
    Sheets("Part Information").Select
    ActiveSheet.Range("$A$1:$Q$3052").AutoFilter Field:=6, Criteria1:= _
        "DDP 2000 T"
    Columns("A:Q").Copy
    Sheets("DDP").Select
    ActiveSheet.Paste
    Sheets("Part Information").Select
    ActiveSheet.Range("$A$1:$Q$3052").AutoFilter Field:=6, Criteria1:=Array( _
        "1500T PRESS", "200T PRESS", "500T PRESS", "800T PRESS"), Operator:= _
        xlFilterValues
    Columns("A:Q").Copy
    Sheets("LightCell").Select
    ActiveSheet.Paste
    Sheets("Part Information").Select
    ActiveSheet.Range("$A$1:$Q$3052").AutoFilter Field:=6, Criteria1:= "HDA OPEN FORGE"
    Columns("A:Q").Copy
    Sheets("Open").Paste
    Columns("A:Q").AutoFilter
    Cells.Select
    Cells.EntireColumn.AutoFit
    Range("A1").Select
    Sheets("LightCell").Select
    Columns("A:Q").AutoFilter
    Cells.Select
    Cells.EntireColumn.AutoFit
    Range("A1").Select
    Sheets("DDP").Select
    Columns("A:Q").AutoFilter
    Cells.Select
    Cells.EntireColumn.AutoFit
    Range("A1").Select
    Sheets("30001000RR").Select
    Columns("A:Q").AutoFilter
    Cells.Select
    Cells.EntireColumn.AutoFit
    Range("A1").Select
    Sheets("25002000T").Select
    Columns("A:Q").AutoFilter
    Cells.Select
    Cells.EntireColumn.AutoFit
    Range("A1").Select
    Sheets("1250T").Select
    Columns("A:Q").AutoFilter
    Cells.Select
    Cells.EntireColumn.AutoFit
    Range("A1").Select
    Sheets("750T").Select
    Columns("A:Q").AutoFilter
    Cells.Select
    Cells.EntireColumn.AutoFit
    Range("A1").Select
    Sheets("12000T").Select
    Columns("D:D").Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Columns("G:G").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("D1").FormulaR1C1 = "Temp"
    Range("G1").FormulaR1C1 = "Setup"
    Range("H1").FormulaR1C1 = "Temp"
    Columns("A:R").AutoFilter
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
