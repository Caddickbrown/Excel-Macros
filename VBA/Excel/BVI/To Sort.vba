'Protecting All Sheets
'Code Stolen and needs cleaning up

Sub protect_all_sheets()
top:
pass = InputBox("password?")
repass = InputBox("Verify Password")
If Not (pass = repass) Then
MsgBox "They didn't match - Try that one again."
GoTo top
End If
For i = 1 To Worksheets.Count
If Worksheets(i).ProtectContents = True Then GoTo oops
Next
For Each s In ActiveWorkbook.Worksheets
s.Protect Password:=pass
Next
Exit Sub
oops: MsgBox "I think you have some sheets that are already protected. Please unprotect all sheets, and then run this Macro."
End Sub

'Below this line I have No Earthly clue what's going on with this code - need to read/understand it.

Sub Move_list_generator()
    ' defines variables
    Dim xworkbook As Workbook, home As Workbook
    Set home = ThisWorkbook
    Dim sht As Worksheet, List As Worksheet, Box_Qty As Worksheet, Required As Worksheet
    Set List = Worksheets("Amco Pick list")
    Set Box_Qty = Worksheets("Box Qty")
    Dim column_count As Integer, boxes As Long, need As Long, delete As Long, yesno As Integer
    Dim SheetName As String
    SheetName = "Kit Schedule Move List" ' <-------- 1) change this if the sheet name of the auto generated move list changes
    Dim located As Boolean

    'identifies size of the current data
    delete = List.UsedRange.Rows.Count

    'deletes all data from the list starting from row 2 until the end as identified above
    Range(List.Cells(2, 1), List.Cells(delete, 5)).Clear

restart: ' <----jump point if it needs to be restarted.

    'searches each open workbook in Excel applicaiton to identify 1 with the predetermined sheet name if this name needs to be updated see above, number 1.
    For Each xworkbook In Application.Workbooks
        If xworkbook.Name <> home.Name Then ' ignores home workbook
            For Each sht In xworkbook.Worksheets 'steps through each sheet in the seleceted workbook.
                If sht.Name = SheetName Then ' compares the name of the sheet against the predetemined name 1
                    If sht.Cells(1, 2) Like ("Kit Schedule Move List*") Then located = True: Exit For 'if correct sheet name checks the cell and checs for the text after the like, this will be the first part of the cell, it does not matter what follows.
                End If
            Next sht
            If located = True Then Exit For ' if found steps out of changing workbook
        End If
    Next xworkbook

    'errors if can't find the sheet required
    If located = False Then
        MsgBox ("Ensure that you have downloaded and opened the automatic kit schedule move list, and enabled the content.") ' If it hasn't found the required sheet, throws an error.
        yesno = MsgBox("Do you want to try again?", vbYesNo) 'asks if the user wants to try again.
        If yesno = 6 Then GoTo restart Else Exit Sub ' vbYesNo gives a numerical outcome, 6 is yes, anyother answer exits the routine
    End If
    'Copies the sheet into the working file, then closes the downloaded workbook
    xworkbook.Worksheets(SheetName).Copy Before:=home.Worksheets("Amco Pick list")
    xworkbook.Close SaveChanges:=False
    Set Required = Sheets(SheetName) ' defines the sheet copied in as required for later calculations

    'finds how many items we need to cycle through
    need = Required.UsedRange.Rows.Count ' need gives how many lines of parts need to be called back
    boxes = Box_Qty.UsedRange.Rows.Count


    'cycles through the sheet of requirements to the end copies across the required data
    For i = 4 To need ' starts at row 4, below the headings
        If Trim(Required.Cells(i, 1)) = "AMCO" Then ' makes sure that only the data from AMCO locations is run, if the location is updated in the future update the text it is searching for.
            For j = 2 To boxes ' For each identified part cycles though all of the parts and box sizes present until it finds the part number
                If Trim(Required.Cells(i, 3)) = Trim(Box_Qty.Cells(j, 1)) Then ' the tirm section removes any additional spaces at hte start / end and ensures everything is a strin whilst comparing, rather than trying to compare a string to number, it doesn't like that
                    List.Cells(i - 2, 1) = Required.Cells(i, 3) ' prints the required part number, i-2 because headings are only 1 line rather than 4.
                    List.Cells(i - 2, 2) = Required.Cells(i, 2) ' prints the required batch
                    List.Cells(i - 2, 3) = Required.Cells(i, 4) 'prints the expiry date
                    List.Cells(i - 2, 4) = Required.Cells(i, 8) ' prints th equantity that is required

                    If Box_Qty.Cells(j, 2) <> 0 Then ' confirms that the box size has been inputted and is greater than 0, if not it states box qty needed.
                        ' If the qty needed is greater than the amount in the location takes everything from the location
                        If Required.Cells(i, 8) >= Required.Cells(i, 7) Then
                            List.Cells(i - 2, 5) = Required.Cells(i, 7)
                        ' if the required parts are in a round number for the box quantity lists the required numbers.
                        ElseIf Required.Cells(i, 8) Mod Box_Qty.Cells(j, 2) = 0 Then List.Cells(i - 2, 5) = Required.Cells(i, 8)
                        'checks if the box quantities required are greater than the quantity in location if so gives locaiton quantity
                        ElseIf WorksheetFunction.RoundUp(Required.Cells(i, 8) / Box_Qty.Cells(j, 2), 0) * Box_Qty.Cells(j, 2) > Required.Cells(i, 7) Then List.Cells(i - 2, 5) = Required.Cells(i, 7)
                        'if box qty * number of boxes is possible states that volume is needed
                        Else: List.Cells(i - 2, 5) = WorksheetFunction.RoundUp(Required.Cells(i, 8) / Box_Qty.Cells(j, 2), 0) * Box_Qty.Cells(j, 2)
                        End If
                    Else: List.Cells(i - 2, 5) = "Box Qty needed" ' states box ty needed if the box qty is blank or = 0
                    End If
                    Exit For
                End If
            Next j
            'if the part number does not appear in the box quantity publishes the parts needed and quantities and states box qty needed
            If List.Cells(i - 2, 1) = "" Then List.Cells(i - 2, 1) = Required.Cells(i, 3): List.Cells(i - 2, 2) = Required.Cells(i, 2): List.Cells(i - 2, 3) = Required.Cells(i, 4): List.Cells(i - 2, 4) = Required.Cells(i, 8): List.Cells(i - 2, 5) = "Box Qty needed"
        End If
    Next i
    'makes the column show as a date
    List.Columns(3).NumberFormat = "dd/mm/yyyy"
    'stops alerts then deletes the requirements
    Application.DisplayAlerts = False
    Required.delete
    Application.DisplayAlerts = True 'starts the alerts again to stop me doing something stupid



End Sub

Sub Record()
'
' Record Macro
'
' Keyboard Shortcut: Ctrl+Shift+L
'

    Dim Total As Integer
    Dim Count As Integer
    Total = 0
    Count = InputBox("How Many Rows?")
    Do While Total < Count
    Selection.EntireRow.Insert , CopyOrigin:=x1FormatFromLeftOrAbove
    Selection.EntireRow.Insert , CopyOrigin:=x1FormatFromLeftOrAbove
    Selection.Offset(3, 0).Select
    Total = Total + 1
    Loop
End Sub

Sub Merging()
'
' Merging Macro
'
' Keyboard Shortcut: Ctrl+Shift+P
'
    Dim Total As Integer
    Dim Count As Integer
    Total = 0
    Count = InputBox("How Many Rows?")
    Do While Total < Count
    Range(ActiveCell.Offset(0, 0), ActiveCell.Offset(2, 0)).Select
    Selection.Merge
    Selection.Offset(1, 0).Select
    Total = Total + 1
    Loop
End Sub

# Refresh Forecast Tables
```
Public Sub RefreshForecastTables()
'******  procedure to refresh all Pivot Tables on a spreadsheet **********
'to be used via a refresh button

Dim sh As Worksheet
Dim pt As PivotTable

'Turn offauto calculate
Application.Calculation = xlManual

'Refreshing Message
Sheets("Refresh").Select

    Range("B10").Select
    Selection.ClearContents
    Range("B9").Select
    Selection.ClearContents
    Range("C10").Select
    Selection.ClearContents
    Range("C9").Select
    Selection.ClearContents

Range("B7:C7").Select
With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 255
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    With Selection.Font
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = 0
    End With

    ActiveCell.Value = " ! REFRESHING ! "

'Turn on auto calculate
Application.Calculation = xlAutomatic

'update queries
Range("B7:C7").Select
   ActiveCell.Value = "UPDATING QUERIES"
ActiveWorkbook.RefreshAll

'update PivotTables
Range("B7:C7").Select
   ActiveCell.Value = "UPDATING PIVOT TABLES"
For Each sh In ActiveWorkbook.Sheets
    For Each pt In sh.PivotTables

        pt.RefreshTable
        DoEvents
        Application.StatusBar = False
    Next
Next

'update queries
Range("B7:C7").Select
   ActiveCell.Value = " ! REFRESHING ! "
ActiveWorkbook.RefreshAll

'Update time last refreshed message
    Sheets("Refresh").Select

   Range("B7:C7").Select
   ActiveCell.Value = "SAVING"

    Range("C9").Select
    ActiveCell.Value = Date
    Range("B9").Select
    ActiveCell.Value = "Sheet Last Refreshed"
    Range("B10").Select
    ActiveCell.Value = "By " & Environ("USERNAME") & " at"
    Range("C10").Select
    ActiveCell.Value = Time()




Range("B7:C7").Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    With Selection.Font
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
    End With
    ActiveCell.Formula = "=IF(C9=TODAY(),""Sheet Has Been Refreshed Today"",""Sheet Needs Refreshing"")"

End Sub
```

Sub Demand_File_Generator()
'
' Demand_File_Generator Macro
'

'
    
    Application.Calculation = xlCalculationManual
    Application.ScreenUpdating = False
    Application.DisplayStatusBar = False
    Application.EnableEvents = False
' Clean Up
    Rows("1:2").Delete
    
    Columns("B:B").Insert Shift:=x1ToRight, CopyOrigin:=x1FormatFromLeftOrAbove
    Columns("B:B").Insert Shift:=x1ToRight, CopyOrigin:=x1FormatFromLeftOrAbove
    Columns("B:B").Insert Shift:=x1ToRight, CopyOrigin:=x1FormatFromLeftOrAbove
    Columns("B:B").Insert Shift:=x1ToRight, CopyOrigin:=x1FormatFromLeftOrAbove
    Columns("G:G").Delete Shift:=xlToLeft
    
' Fill out column names
    Range("B1").FormulaR1C1 = "Dist"
    Range("C1").FormulaR1C1 = "DW?"
    Range("D1").FormulaR1C1 = "Priority"
    Range("E1").FormulaR1C1 = "Country"
    Range("L1").FormulaR1C1 = "Picks"
    Range("M1").FormulaR1C1 = "CPU"
    Range("N1").FormulaR1C1 = "Boxes"
    Columns("J:K").Cut
    Columns("O:O").Insert Shift:=xlToRight

    Range("O1").FormulaR1C1 = "Constraint"
    Range("P1").FormulaR1C1 = "Notes"
    Range("Q1").FormulaR1C1 = "Action"
    Range("R1").FormulaR1C1 = "BOM Check"
' Dist Column
    Range("B2").FormulaR1C1 = _
        "=VLOOKUP(RC[-1],'https://bvx.sharepoint.com/Operations/Bidford/[KIT Substitutions Master List (incl item swap list)- 14JAN2022.xlsx]Direct- Indirect PACK TYPE'!C1:C5,5,0)"
' Double Wraps
    Range("C2").FormulaR1C1 = _
        "=HLOOKUP(RC[-2],'S:\Public\Kit Standard Times\[Std Time New Format.xlsm]Kit Data'!R3:R11,9,FALSE)"
' Country
    Range("E2").FormulaR1C1 = _
        "=VLOOKUP(RC[-4],'https://bvx.sharepoint.com/Operations/Bidford/[KIT Substitutions Master List (incl item swap list)- 14JAN2022.xlsx]Direct- Indirect PACK TYPE'!C1:C5,3,0)"
' Picks
    Range("J2").FormulaR1C1 = _
        "=HLOOKUP(RC[-9],'S:\Public\Kit Standard Times\[Std Time New Format.xlsm]Kit Data'!R3:R8,6,FALSE)"
' CPU
    Range("K2").FormulaR1C1 = _
        "=IFERROR(HLOOKUP(RC[-10],'S:\Public\Kit Standard Times\[Std Time New Format.xlsm]Kit Data'!R3:R11,4,FALSE),"""")"
' Boxes
    Range("L2").FormulaR1C1 = _
       "=IFERROR(IF(RC[-9]=" & Chr(34) & "DC" & Chr(34) & ",ROUNDUP(RC[-3]/RC[-1]/10,0),RC[-3]/RC[-1]),RC[-3]/LEFT(RC[-1],2))"
    
    Range("B2:E2").AutoFill Destination:=Range("B2:E8000")

    Range("J2:L2").AutoFill Destination:=Range("J2:L8000")

    Columns("A:A").TextToColumns Destination:=Range("A1"), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=False, _
        Semicolon:=False, Comma:=False, Space:=False, Other:=False, FieldInfo _
        :=Array(1, 1), TrailingMinusNumbers:=True
    Range("O2").FormulaR1C1 = "=IF(AND(RC[-2]="""",RC[-1]=""""),""N"",""Y"")"
    Range("O2").AutoFill Destination:=Range("O2:O1667")
    Range("R2").AutoFill Destination:=Range("R2:R1667")
    Columns("A:R").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Range("A1:R1").Font.Bold = True

    Cells.EntireColumn.AutoFit
    Range("I1").FormulaR1C1 = "Qty"
    Columns("I:I").EntireColumn.AutoFit
    Range("A1:R1").Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent5
        .TintAndShade = -0.249977111117893
        .PatternTintAndShade = 0
    End With

    Range("A1").Select
    Application.EnableEvents = True
    Application.DisplayStatusBar = True
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic

End Sub


Sub Record()
'
' Record Macro
'
' Keyboard Shortcut: Ctrl+Shift+L
'
   
    Dim Total As Integer
    Dim Count As Integer
    Total = 0
    Count = InputBox("How Many Rows?")
    Do While Total < Count
    Selection.EntireRow.Insert , CopyOrigin:=x1FormatFromLeftOrAbove
    Selection.EntireRow.Insert , CopyOrigin:=x1FormatFromLeftOrAbove
    Selection.Offset(3, 0).Select
    Total = Total + 1
    Loop
End Sub

Sub Merging()
'
' Merging Macro
'
' Keyboard Shortcut: Ctrl+Shift+P
'
    Dim Total As Integer
    Dim Count As Integer
    Total = 0
    Count = InputBox("How Many Rows?")
    Do While Total < Count
    Range(ActiveCell.Offset(0, 0), ActiveCell.Offset(2, 0)).Select
    Selection.Merge
    Selection.Offset(1, 0).Select
    Total = Total + 1
    Loop
End Sub


Sub RenameSheet()

    Dim rs As Worksheet

    For Each rs In Sheets
    rs.Name = rs.Range("B2")
    Next rs

End Sub


Sub TTC()
'
' TTC Macro
'

'
    Columns("A:A").Select
    Application.CutCopyMode = False
    Selection.TextToColumns Destination:=Range("A1"), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=False, _
        Semicolon:=False, Comma:=False, Space:=False, Other:=False, FieldInfo _
        :=Array(1, 1), TrailingMinusNumbers:=True
    Columns("B:B").Select
    Application.CutCopyMode = False
    Selection.TextToColumns Destination:=Range("A1"), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=False, _
        Semicolon:=False, Comma:=False, Space:=False, Other:=False, FieldInfo _
        :=Array(1, 1), TrailingMinusNumbers:=True
    Columns("C:C").Select
    Application.CutCopyMode = False
    Selection.TextToColumns Destination:=Range("A1"), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=False, _
        Semicolon:=False, Comma:=False, Space:=False, Other:=False, FieldInfo _
        :=Array(1, 1), TrailingMinusNumbers:=True
    Columns("D:D").Select
    Application.CutCopyMode = False
    Selection.TextToColumns Destination:=Range("A1"), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=False, _
        Semicolon:=False, Comma:=False, Space:=False, Other:=False, FieldInfo _
        :=Array(1, 1), TrailingMinusNumbers:=True
    Columns("E:E").Select
    Application.CutCopyMode = False
    Selection.TextToColumns Destination:=Range("A1"), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=False, _
        Semicolon:=False, Comma:=False, Space:=False, Other:=False, FieldInfo _
        :=Array(1, 1), TrailingMinusNumbers:=True
    Columns("F:F").Select
    Application.CutCopyMode = False
    Selection.TextToColumns Destination:=Range("A1"), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=False, _
        Semicolon:=False, Comma:=False, Space:=False, Other:=False, FieldInfo _
        :=Array(1, 1), TrailingMinusNumbers:=True
    Columns("G:G").Select
    Application.CutCopyMode = False
    Selection.TextToColumns Destination:=Range("A1"), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=False, _
        Semicolon:=False, Comma:=False, Space:=False, Other:=False, FieldInfo _
        :=Array(1, 1), TrailingMinusNumbers:=True
    Columns("H:H").Select
    Application.CutCopyMode = False
    Selection.TextToColumns Destination:=Range("A1"), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=False, _
        Semicolon:=False, Comma:=False, Space:=False, Other:=False, FieldInfo _
        :=Array(1, 1), TrailingMinusNumbers:=True
    Columns("I:I").Select
    Application.CutCopyMode = False
    Selection.TextToColumns Destination:=Range("A1"), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=False, _
        Semicolon:=False, Comma:=False, Space:=False, Other:=False, FieldInfo _
        :=Array(1, 1), TrailingMinusNumbers:=True
End Sub

Sub TEXT_TO_COLUMN()
'
' TEXT_TO_COLUMN Macro
'

'
    Columns("B:B").Select
    Selection.TextToColumns Destination:=Range("B1"), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=False, _
        Semicolon:=False, Comma:=False, Space:=False, Other:=False, FieldInfo _
        :=Array(1, 1), TrailingMinusNumbers:=True
    Range("J12").Select
End Sub

Sub RFQPlanPOs()
'
' RFQPlanPOs Macro
'

'
    Application.Calculation = xlCalculationManual
    Application.ScreenUpdating = False
    Application.DisplayStatusBar = False
    Application.EnableEvents = False
    
    Rows("1:1").AutoFilter
    Range("A:J,L:L,N:U,W:FI").Delete Shift:=xlToLeft
    Columns("A:C").EntireColumn.AutoFit
    Columns("A:C").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Range("A1:C1").Font.Bold = True
    Range("A1").Select
    
    Application.EnableEvents = True
    Application.DisplayStatusBar = True
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    
End Sub

Sub RFQPlanMRP()
'
' RFQPlanMRP Macro
'

'
    Application.Calculation = xlCalculationManual
    Application.ScreenUpdating = False
    Application.DisplayStatusBar = False
    Application.EnableEvents = False
    
    Rows("1:1").AutoFilter
    Range("D:J").Delete Shift:=xlToLeft
    Columns("A:C").EntireColumn.AutoFit
    Columns("A:C").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Range("A1:C1").Font.Bold = True
    Range("A1").Select
    
    Application.EnableEvents = True
    Application.DisplayStatusBar = True
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    
End Sub

Sub RFQPlanStock()
'
' RFQPlanStock Macro
'

'
    Application.Calculation = xlCalculationManual
    Application.ScreenUpdating = False
    Application.DisplayStatusBar = False
    Application.EnableEvents = False
    
    Rows("1:1").AutoFilter
    Range("D:D,F:R,T:BH").Delete Shift:=xlToLeft
    Columns("A:E").EntireColumn.AutoFit
    Columns("A:E").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Range("A1:E1").Font.Bold = True
    Range("A1").Select
    
    Application.EnableEvents = True
    Application.DisplayStatusBar = True
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    
End Sub
