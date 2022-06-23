'These are various Macros to Control a "Next Up" Sheet in Excel. They were assigned to buttons to run actions from the one screen. Generally they move things along a "Route" in the different "Operations".

'WarehouseDone - Looks up a value, finds it in a different sheet and changes a value that says "DHR" on the same line to "Warehouse"
'WarehouseUndo - Looks up a value, finds it in a different sheet and changes it "back" to "DHR" (which is what Warehouse would look up to)
'PrekitDone - Looks up a value, finds it in a different sheet and changes a value that says "Warehouse" on the same line to "Prekit"
'PrekitUndo - Looks up a value, finds it in a different sheet and changes it "back" to "Warehouse" (which is what Prekit would look up to)
'OnLineUndo - Looks up a value, finds it in a different sheet and changes it "back" to "Prekit" (which is what On Line would look up to)
'OnHold - Looks up a value, finds it in a different sheet and changes it to "ON HOLD"
'Complete -  Looks up a value, finds it in a different sheet and changes it to "Completed"

Sub WarehouseDone()
Application.Calculation = xlCalculationManual
Application.ScreenUpdating = False
Application.DisplayStatusBar = False
Application.EnableEvents = False

Temptext = Worksheets("NextUp").Range("C2")
    Sheets("Main").Select
    Cells.Find(What:=Temptext, After:=ActiveCell, LookIn:=xlValues, LookAt _
        :=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:= _
        False, SearchFormat:=False).Activate
    ActiveCell.Offset(0, 14).Replace What:="DHR", Replacement:="Warehouse", LookAt:=xlPart _
        , SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Sheets("NextUp").Select

Application.EnableEvents = True
Application.DisplayStatusBar = True
Application.ScreenUpdating = True
Application.Calculation = xlCalculationAutomatic
End Sub

Sub WarehouseUndo()
Application.Calculation = xlCalculationManual
Application.ScreenUpdating = False
Application.DisplayStatusBar = False
Application.EnableEvents = False

Temptext = InputBox("What Shop Order needs bringing back to the Warehouse?")
    Sheets("Main").Select
    Cells.Find(What:=Temptext, After:=ActiveCell, LookIn:=xlValues, LookAt _
        :=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:= _
        False, SearchFormat:=False).Activate
    ActiveCell.Offset(0, 14).Replace What:="Warehouse", Replacement:="DHR", LookAt:=xlPart _
        , SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    ActiveCell.Offset(0, 14).Replace What:="Prekit", Replacement:="DHR", LookAt:=xlPart _
        , SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    ActiveCell.Offset(0, 14).Replace What:="On Line", Replacement:="DHR", LookAt:=xlPart _
        , SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    ActiveCell.Offset(0, 14).Replace What:="Completed", Replacement:="DHR", LookAt:=xlPart _
        , SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    ActiveCell.Offset(0, 14).Replace What:="ON HOLD", Replacement:="DHR", LookAt:=xlPart _
        , SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Sheets("NextUp").Select

Application.EnableEvents = True
Application.DisplayStatusBar = True
Application.ScreenUpdating = True
Application.Calculation = xlCalculationAutomatic
End Sub

Sub PrekitDone()
Application.Calculation = xlCalculationManual
Application.ScreenUpdating = False
Application.DisplayStatusBar = False
Application.EnableEvents = False

Temptext = Worksheets("NextUp").Range("C3")
    Sheets("Main").Select
    Cells.Find(What:=Temptext, After:=ActiveCell, LookIn:=xlValues, LookAt _
        :=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:= _
        False, SearchFormat:=False).Activate
    ActiveCell.Offset(0, 14).Replace What:="Warehouse", Replacement:="Prekit", LookAt:=xlPart _
        , SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Sheets("NextUp").Select

Application.EnableEvents = True
Application.DisplayStatusBar = True
Application.ScreenUpdating = True
Application.Calculation = xlCalculationAutomatic
End Sub

Sub PrekitUndo()
Application.Calculation = xlCalculationManual
Application.ScreenUpdating = False
Application.DisplayStatusBar = False
Application.EnableEvents = False

Temptext = InputBox("What Shop Order needs bringing back to Prekit?")
    Sheets("Main").Select
    Cells.Find(What:=Temptext, After:=ActiveCell, LookIn:=xlValues, LookAt _
        :=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:= _
        False, SearchFormat:=False).Activate
    ActiveCell.Offset(0, 14).Replace What:="Prekit", Replacement:="Warehouse", LookAt:=xlPart _
        , SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    ActiveCell.Offset(0, 14).Replace What:="On Line", Replacement:="Warehouse", LookAt:=xlPart _
        , SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    ActiveCell.Offset(0, 14).Replace What:="Completed", Replacement:="Warehouse", LookAt:=xlPart _
        , SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    ActiveCell.Offset(0, 14).Replace What:="ON HOLD", Replacement:="Warehouse", LookAt:=xlPart _
        , SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Sheets("NextUp").Select

Application.EnableEvents = True
Application.DisplayStatusBar = True
Application.ScreenUpdating = True
Application.Calculation = xlCalculationAutomatic
End Sub

Sub OnLineUndo()
Application.Calculation = xlCalculationManual
Application.ScreenUpdating = False
Application.DisplayStatusBar = False
Application.EnableEvents = False

Temptext = InputBox("What Shop Order needs bringing back to the Line?")
    Sheets("Main").Select
    Cells.Find(What:=Temptext, After:=ActiveCell, LookIn:=xlValues, LookAt _
        :=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:= _
        False, SearchFormat:=False).Activate
    ActiveCell.Offset(0, 14).Replace What:="Warehouse", Replacement:="Prekit", LookAt:=xlPart _
        , SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    ActiveCell.Offset(0, 14).Replace What:="On Line", Replacement:="Prekit", LookAt:=xlPart _
        , SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    ActiveCell.Offset(0, 14).Replace What:="Completed", Replacement:="Prekit", LookAt:=xlPart _
        , SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    ActiveCell.Offset(0, 14).Replace What:="ON HOLD", Replacement:="Prekit", LookAt:=xlPart _
        , SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Sheets("NextUp").Select

Application.EnableEvents = True
Application.DisplayStatusBar = True
Application.ScreenUpdating = True
Application.Calculation = xlCalculationAutomatic
End Sub

Sub OnHold()
Application.Calculation = xlCalculationManual
Application.ScreenUpdating = False
Application.DisplayStatusBar = False
Application.EnableEvents = False

Temptext = InputBox("What Shop Order needs putting On Hold?")
Tempreason = InputBox("What has gone wrong?")
    Sheets("Main").Select
    Cells.Find(What:=Temptext, After:=ActiveCell, LookIn:=xlValues, LookAt _
        :=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:= _
        False, SearchFormat:=False).Activate
    ActiveCell.Offset(0, 14).Replace What:="DHR", Replacement:="ON HOLD", LookAt:=xlPart _
        , SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    ActiveCell.Offset(0, 14).Replace What:="Warehouse", Replacement:="ON HOLD", LookAt:=xlPart _
        , SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    ActiveCell.Offset(0, 14).Replace What:="Prekit", Replacement:="ON HOLD", LookAt:=xlPart _
        , SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    ActiveCell.Offset(0, 14).Replace What:="On Line", Replacement:="ON HOLD", LookAt:=xlPart _
        , SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    ActiveCell.Offset(0, 17) = Tempreason
    Sheets("NextUp").Select

Application.EnableEvents = True
Application.DisplayStatusBar = True
Application.ScreenUpdating = True
Application.Calculation = xlCalculationAutomatic
End Sub

Sub Complete()
Application.Calculation = xlCalculationManual
Application.ScreenUpdating = False
Application.DisplayStatusBar = False
Application.EnableEvents = False

Temptext = InputBox("What Shop Order can be Completed?")
    Sheets("Main").Select
    Cells.Find(What:=Temptext, After:=ActiveCell, LookIn:=xlValues, LookAt _
        :=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:= _
        False, SearchFormat:=False).Activate
    
    ActiveCell.Offset(0, 14).Replace What:="DHR", Replacement:="Completed", LookAt:=xlPart _
        , SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    ActiveCell.Offset(0, 14).Replace What:="Warehouse", Replacement:="Completed", LookAt:=xlPart _
        , SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    ActiveCell.Offset(0, 14).Replace What:="Prekit", Replacement:="Completed", LookAt:=xlPart _
        , SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    ActiveCell.Offset(0, 14).Replace What:="On Line", Replacement:="Completed", LookAt:=xlPart _
        , SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    ActiveCell.Offset(0, 14).Replace What:="ON HOLD", Replacement:="Completed", LookAt:=xlPart _
        , SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2

    Sheets("NextUp").Select

Application.EnableEvents = True
Application.DisplayStatusBar = True
Application.ScreenUpdating = True
Application.Calculation = xlCalculationAutomatic
End Sub