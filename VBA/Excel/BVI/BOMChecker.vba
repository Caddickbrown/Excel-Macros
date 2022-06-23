'The Below are two Macros to make a "BOM Checker Sheet"

'NewPart allows to to change the Parent Part Number/Qty to Lookup without the ability to delete the part number making the "Filter" Formula crash Excel and start trying to lookup any blank lines. It then drops down the formulas so as to reserve memory.

'ClearOut clears the Part Number and Qty Safely.

Sub NewPart()

Application.Calculation = xlCalculationManual
Application.ScreenUpdating = False
Application.DisplayStatusBar = False
Application.EnableEvents = False

Dim last_row As Long
    
    Range("B1:B2") = "-"
    Range("C6:H9000").ClearContents
    
    TempPart = InputBox("What is the Part Number?")
    
    If (TempPart = "") Then
    
        MsgBox "Missing Part Number", , "Error"
            
        Application.EnableEvents = True
        Application.DisplayStatusBar = True
        Application.ScreenUpdating = True
        Application.Calculation = xlCalculationAutomatic
        
        Exit Sub
    
    Else

        TempQty = InputBox("What Quantity do you want to check?")
    
        If (TempQty = "") Then
    
            MsgBox "Missing Quantity", , "Error"

            Application.EnableEvents = True
            Application.DisplayStatusBar = True
            Application.ScreenUpdating = True
            Application.Calculation = xlCalculationAutomatic
           
            Exit Sub
        
        Else
    
            Application.Calculation = xlCalculationAutomatic
            Range("B1") = TempPart
            Range("B2") = TempQty

            last_row = Cells(Rows.Count, 1).End(xlUp).Row
            Range("C5:H5").AutoFill Destination:=Range("C5:H" & last_row)

            Columns("B:B").EntireColumn.AutoFit

        End If

    End If

Application.EnableEvents = True
Application.DisplayStatusBar = True
Application.ScreenUpdating = True
Application.Calculation = xlCalculationAutomatic

End Sub


Sub ClearOut()

Application.Calculation = xlCalculationManual
Application.ScreenUpdating = False
Application.DisplayStatusBar = False
Application.EnableEvents = False

    Range("B1:B2") = "-"
    Range("C6:H9000").ClearContents

Application.EnableEvents = True
Application.DisplayStatusBar = True
Application.ScreenUpdating = True
Application.Calculation = xlCalculationAutomatic

End Sub
