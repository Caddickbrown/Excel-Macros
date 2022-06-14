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
    SheetName = "Kit Schedule Move List" ' <-------- 1) change this if the sheet name of tyhe auto generated move list changes
    Dim located As Boolean

    'identifies size of the current data
    delete = List.UsedRange.Rows.Count

    'deletes all data from teh list starting from row 2 until the end as identified above
    Range(List.Cells(2, 1), List.Cells(delete, 5)).Clear

restart: ' <----jump point if it needs to be restarted.

    'searches each open workbook in Excel applicaiton to identify 1 with the predetermined sheet name if this name needs to be updated see above, number 1.
    For Each xworkbook In Application.Workbooks
        If xworkbook.Name <> home.Name Then ' ignores home workbook
            For Each sht In xworkbook.Worksheets 'steps through each sheet in the seleceted workbook.
                If sht.Name = SheetName Then ' compares teh name of the sheet against the predetemined name 1
                    If sht.Cells(1, 2) Like ("Kit Schedule Move List*") Then located = True: Exit For 'if correct sheet name checks the cell and checs for the text after the like, this will be the first part of the cell, it doesnot matter what follows.
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
    need = Required.UsedRange.Rows.Count ' need gives howmany lines of parts need to be called back
    boxes = Box_Qty.UsedRange.Rows.Count


    'cycles through the sheet of requirements to the end copies across teh required data
    For i = 4 To need ' starts at row 4, below the headings
        If Trim(Required.Cells(i, 1)) = "AMCO" Then ' makes sure that only the data from AMCO locations is run, if the location is updated in teh future update the text it is searching for.
            For j = 2 To boxes ' For each identified part cycles though all of the parts and box sizes present until it finds the part number
                If Trim(Required.Cells(i, 3)) = Trim(Box_Qty.Cells(j, 1)) Then ' the tirm section removes any additional spaces at hte start / end and ensures everything is a strin whilst comparing, rather than trying to compare a string to number, it doesn't like that
                    List.Cells(i - 2, 1) = Required.Cells(i, 3) ' prints the required part number, i-2 because headings are only 1 line rather than 4.
                    List.Cells(i - 2, 2) = Required.Cells(i, 2) ' prints the required batch
                    List.Cells(i - 2, 3) = Required.Cells(i, 4) 'prints teh expiry date
                    List.Cells(i - 2, 4) = Required.Cells(i, 8) ' prints th equantity that is required

                    If Box_Qty.Cells(j, 2) <> 0 Then ' confirms that the box size has been inputted and is greater than 0, if not it states box qty needed.
                        ' If the qty needed is greater than the amount in teh location takes everything from the location
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
            'if the part number does not appear in teh box quantity publishes the parts needed and quantities and states box qty needed
            If List.Cells(i - 2, 1) = "" Then List.Cells(i - 2, 1) = Required.Cells(i, 3): List.Cells(i - 2, 2) = Required.Cells(i, 2): List.Cells(i - 2, 3) = Required.Cells(i, 4): List.Cells(i - 2, 4) = Required.Cells(i, 8): List.Cells(i - 2, 5) = "Box Qty needed"
        End If
    Next i
    'makes the column show as a date
    List.Columns(3).NumberFormat = "dd/mm/yyyy"
    ' stops alerts then deletes the requirements
    Application.DisplayAlerts = False
    Required.delete
    Application.DisplayAlerts = True ' starts the alerts again to stop me doing something stupid



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
