'Protecting All Sheets
'Code Stolen and needs cleaning up

Sub protect_all_sheets()
top:
pass = InputBox("password?")
repass = InputBox("Verify Password")
If Not (pass = repass) Then
MsgBox "you made a boo boo"
GoTo top
End If
For i = 1 To Worksheets.Count
If Worksheets(i).ProtectContents = True Then GoTo oops
Next
For Each s In ActiveWorkbook.Worksheets
s.Protect Password:=pass
Next
Exit Sub
oops: MsgBox "I think you have some sheets that are already protected. Please unprotect all sheets then running this Macro."
End Sub

'Below this line I have No Earthly clue what's going on with this code - need to read/understand it.

Sub SIC()
    Dim Home As Workbook
    Set Home = ThisWorkbook
    Dim temp As Worksheet
    Set temp = Worksheets("Template")
    Dim SIC As Worksheet
    Dim data As Worksheet

    Dim xWorkbook As Workbook
    Dim sht As Worksheet
    Dim Today As Date
    Today = Date
    'Today = #11/25/2021#
    Dim yesterday As Date
    yesterday = Today - 1
    Dim test As String
    test = Format(Today, "ddmmmyy")
    Dim SheetName As String
    SheetName = "OverviewInventoryTransactionHis"
    Dim colum_count As Integer
    Dim i As Integer
    Dim j As Integer
    Dim k As Integer
    Dim bay As Integer
    Dim rowcount As Integer
    Dim CreateTime As Integer
    Dim Create As Integer
    Dim clock
    Dim LastComp
    Dim Picks As Integer
    Dim Picker() As String
    Dim pickers As Integer
    Dim outputrow As Integer
    Dim count As Integer
    Dim N_pick As Integer
    Dim N_Pickers As Double
    Dim M_Pick As Integer
    Dim M_Pickers As Double
    Dim A_pick As Integer
    Dim A_Pickers As Double
    Dim previous As String
    Dim Short As Integer
    Dim Date_present As Boolean

    ' checks if workbook is read only
    If Home.ReadOnly Then
        i = MsgBox("SIC sheet is opened as Read Only. Please reopen the file and ensure that it is not Read Only before trying to run. Do you wisjh to continue?", vbYesNo)
        If i = 6 Then MsgBox ("Please save file under a different name."): Exit Sub
        If i = 7 Then Home.Close
    End If
    ' checks if current day has a sheet
    For Each sht In Application.ThisWorkbook.Worksheets
        If sht.Name = test Then Set SIC = sht: GoTo jump
    Next sht
    Worksheets("template").Copy After:=Sheets(Sheets.count)
    Set SIC = Sheets(Sheets.count)
    SIC.Name = test
    SIC.Cells(1, 13) = Today
jump:
   Call Archive
    ' looks through each workbook that is open
    For Each xWorkbook In Application.Workbooks
        If xWorkbook.Name <> Home.Name Then ' ignores home workbook
            For Each sht In xWorkbook.Worksheets 'looks at each sheet in current workbook
                If sht.Name = SheetName Then
                    column_count = sht.Cells(1, Columns.count).End(xlToLeft).Column
                    For i = 1 To column_count
                        If sht.Cells(1, i) = "Bay" Then bay = i
                        If sht.Cells(1, i) = "Created" Then Create = i
                        If bay > 0 And Create > 0 Then Exit For ' searches row 1 for the column containing the required data if present exits loop having recorded column
                    Next i
                    If sht.Cells(2, Create) = Today Or sht.Cells(2, Create) = yesterday Then
                        Date_present = True
                        If sht.Cells(2, bay) = "SOM" Or sht.Cells(2, bay) = "MSOM" Or sht.Cells(2, bay) = "PK" Then located = True: Exit For ' searches the identified row for what needs the required info
                    Else: bay = 0: Create = 0
                    End If
                End If
            Next sht
            If located = True Then Exit For
        End If
    Next xWorkbook
    'cancels out of the program if there is no open data file.
    If Date_present = False Then MsgBox ("There is no data open from the correct dates, the data should be from " & yesterday & " or " & Today): Exit Sub
    If located = False Then MsgBox ("You must download the data from IFS then rerun the program"): Exit Sub

    xWorkbook.Sheets(SheetName).Copy Before:=Home.Worksheets("Targets")
    Set data = Sheets(SheetName)

    xWorkbook.Close SaveChanges:=False

    test = Format(data.Cells(2, Create), "ddmmmyy")
    Set SIC = Sheets(test)
    SIC.Name = test

    column_count = data.Cells(1, Columns.count).End(xlToLeft).Column
    'finds time& performed by column
    For i = 1 To column_count
        If data.Cells(1, i) = "Creation Time" Then CreateTime = i
        If data.Cells(1, i) = "Performed By" Then pickers = i
        If CreateTime > 0 And pickers > 0 Then Exit For
    Next i
    'sorts by time
    rowcount = data.Cells(Rows.count, CreateTime).End(xlUp).Row
    Range(data.Cells(1, 1), data.Cells(rowcount, column_count)).Sort Key1:=Range(data.Cells(1, CreateTime), data.Cells(rowcount, CreateTime)), Order1:=xlAscending, Header:=xlYes
    j = 0
    located = False
    'Last complete hour
    clock = Time
    ReDim Picker(0)
    If data.Cells(2, Create) = Today Then count = CInt(Hour(clock)) Else If data.Cells(2, Create) = yesterday Then count = 24
    For i = Hour(SIC.Cells(8, 14)) + 1 To count
        If j = 0 Then k = 2 Else k = j
        For j = k To rowcount
            If Hour(data.Cells(j, CreateTime)) > i - 2 Then
            If Hour(data.Cells(j, CreateTime)) < i Then
                pick = pick + 1
                If data.Cells(j, bay) = "PK" Then Short = Short + 1
                For k = 0 To UBound(Picker())
                    If data.Cells(j, pickers) = Picker(k) Then GoTo present
                Next k
                ReDim Preserve Picker(UBound(Picker()) + 1)
                Picker(UBound(Picker())) = data.Cells(j, pickers)
present:
            Else: Exit For
            End If
            End If
        Next j
        SIC.Cells(i + 2, 11) = Sheets("Targets").Cells(6, 2)
        SIC.Cells(i + 2, 2) = pick
        SIC.Cells(i + 2, 4) = UBound(Picker())
        If i = 2 Or i = 5 Or i = 10 Or i = 13 Or i = 18 Or i = 21 Then SIC.Cells(i + 2, 5) = Sheets("Targets").Cells(2, 2) * 0.75 Else SIC.Cells(i + 2, 5) = Sheets("Targets").Cells(2, 2)
        If SIC.Cells(i + 2, 4) > 0 Then SIC.Cells(i + 2, 6) = Round(pick / SIC.Cells(i + 2, 4), 2) Else SIC.Cells(i + 2, 6) = 0
        If SIC.Cells(i + 2, 6) < SIC.Cells(i + 2, 5) And SIC.Cells(i + 2, 6) > 0 Then SIC.Cells(i + 2, 6).Interior.ColorIndex = 3 Else If SIC.Cells(i + 2, 6) > 0 Then SIC.Cells(i + 2, 6).Interior.ColorIndex = 4
        SIC.Cells(i + 2, 7) = Short
        pick = 0
        Short = 0
        ReDim Picker(0)
    Next i

    LastComp = TimeSerial(i - 1, 0, 0)
    SIC.Cells(8, 14) = LastComp

    Application.DisplayAlerts = False 'turns off pop-ups to ask if ok to delete sheet
    Sheets(SheetName).Delete 'Deletes sheet
    Application.DisplayAlerts = True 'turns pop-ups back on to stop me doing something silly

    previous = Format(SIC.Cells(1, 13) - 1, "ddmmmyy")

    N_pick = Worksheets(previous).Cells(25, 2) + Worksheets(previous).Cells(26, 2)
    N_Pickers = Worksheets(previous).Cells(25, 4) + Worksheets(previous).Cells(26, 4)
    For i = 3 To SIC.Cells(Rows.count, 4).End(xlUp).Row
        If i <= 8 Then N_pick = N_pick + SIC.Cells(i, 2): If i = 4 Or i = 7 Then N_Pickers = N_Pickers + SIC.Cells(i, 4) * 0.75 Else N_Pickers = N_Pickers + SIC.Cells(i, 4)
        If i <= 16 And i > 8 Then M_Pick = M_Pick + SIC.Cells(i, 2): If i = 12 Or i = 15 Then M_Pickers = M_Pickers + SIC.Cells(i, 4) * 0.75 Else M_Pickers = M_Pickers + SIC.Cells(i, 4)
        If i <= 24 And i > 16 Then A_pick = A_pick + SIC.Cells(i, 2): If i = 20 Or i = 23 Then A_Pickers = A_Pickers + SIC.Cells(i, 4) * 0.75 Else A_Pickers = A_Pickers + SIC.Cells(i, 4)
    Next i
    SIC.Cells(12, 13) = N_pick
    SIC.Cells(12, 14) = N_Pickers
    If SIC.Cells(12, 14) > 0 Then SIC.Cells(12, 15) = Round(SIC.Cells(12, 13) / SIC.Cells(12, 14), 2)
    If SIC.Cells(12, 15) < Worksheets("Targets").Cells(2, 2) And SIC.Cells(12, 15) > 0 Then SIC.Cells(12, 15).Interior.ColorIndex = 3 Else If SIC.Cells(12, 15) > 0 Then SIC.Cells(12, 15).Interior.ColorIndex = 4
    SIC.Cells(13, 13) = M_Pick
    SIC.Cells(13, 14) = M_Pickers
    If SIC.Cells(13, 14) > 0 Then SIC.Cells(13, 15) = Round(SIC.Cells(13, 13) / SIC.Cells(13, 14), 2)
    If SIC.Cells(13, 15) < Sheets("Targets").Cells(2, 2) And SIC.Cells(13, 15) > 0 Then SIC.Cells(13, 15).Interior.ColorIndex = 3 Else If SIC.Cells(13, 15) > 0 Then SIC.Cells(13, 15).Interior.ColorIndex = 4
    SIC.Cells(14, 13) = A_pick
    SIC.Cells(14, 14) = A_Pickers
    If SIC.Cells(14, 14) > 0 Then SIC.Cells(14, 15) = Round(SIC.Cells(14, 13) / SIC.Cells(14, 14), 2)
    If SIC.Cells(14, 15) < Sheets("Targets").Cells(2, 2) And SIC.Cells(14, 15) > 0 Then SIC.Cells(14, 15).Interior.ColorIndex = 3 Else If SIC.Cells(14, 15) > 0 Then SIC.Cells(14, 15).Interior.ColorIndex = 4
    SIC.Cells(15, 13) = SIC.Cells(12, 13) + SIC.Cells(13, 13) + SIC.Cells(14, 13)
    SIC.Cells(15, 14) = SIC.Cells(12, 14) + SIC.Cells(13, 14) + SIC.Cells(14, 14)
    If SIC.Cells(15, 14) > 0 Then SIC.Cells(15, 15) = Round(SIC.Cells(15, 13) / SIC.Cells(15, 14), 2)
    If SIC.Cells(15, 15) < Sheets("Targets").Cells(2, 2) And SIC.Cells(15, 15) > 0 Then SIC.Cells(15, 15).Interior.ColorIndex = 3 Else If SIC.Cells(15, 15) > 0 Then SIC.Cells(15, 15).Interior.ColorIndex = 4

    SIC.Select
    Home.Save

End Sub

Sub Archive()
    Dim Home As Workbook
    Set Home = ThisWorkbook

    Dim count As Integer
    count = Sheets.count ' count = number of sheets in workbook
    Dim sht As Worksheet
    Dim oldest As Date
    Dim newest As Date
    Dim days As Integer
    Dim folder As String
    folder = ThisWorkbook.Path ' as long as archive and final SIC sheet are stored int eh same folder this will work.
    Dim file As String
    file = "SIC_ARCHIVE.xlsm" ' file to store data in
    Dim Archive As Workbook
    Dim filepath As String
    filepath = folder & "\" & file
    'Workbooks.Open (filepath)
    Dim present As Boolean

    'Set Archive = Workbooks(file)

    If count > 8 Then

        For Each sht In Home.Worksheets
            If sht.Name Like ("Sheet*") Then ' searches and deletes anything that is titled SheetX
                If Application.WorksheetFunction.CountA(sht.Cells) = 0 Then Application.DisplayAlerts = False: sht.Delete: Application.DisplayAlerts = True
                If Sheets.count <= 8 Then Exit Sub ' if still more than 8 sheets continues
            End If
            If sht.Name Like ("##***##") Then ' records number of days worth of production and start and finish date
                days = days + 1
                If oldest = #12:00:00 AM# Then oldest = sht.Cells(1, 13) Else If oldest > sht.Cells(1, 13) Then oldest = sht.Cells(1, 13)
                If newest = #12:00:00 AM# Then newest = sht.Cells(1, 13) Else If newest < sht.Cells(1, 13) Then newest = sht.Cells(1, 13)

            End If

        Next sht
        If days > 5 Then ' if more than 5 days or production move oldest SIC sheets to archive file
        Workbooks.Open (filepath) ' opens the workbook as defined above
        Set Archive = Workbooks(file) ' stores it as archive workboook
        If Archive.ReadOnly = True Then Exit Sub
move:
            Home.Worksheets(Format(oldest, "ddmmmyy")).move After:=Archive.Worksheets(Worksheets.count) 'moves the oldest sheet to archive
            days = days - 1 ' subtracts 1 from day count
            If days > 5 Then oldest = oldest + 1: Call exists(oldest, newest, Home, present): If present = True Then GoTo move 'if still more than 5 days worth of production it will find the next oldest date and move the sheet to archive
            End If
        Archive.Save
        Archive.Close
    End If
    'Archive.Save
    'Archive.Close
End Sub

Function exists(old As Date, newest As Date, Home As Workbook, present As Boolean)

    Dim sht As Worksheet
    present = False
redo:
    For Each sht In Home.Worksheets
        If sht.Name = Format(old, "ddmmmyy") Then present = True: Exit Function
    Next sht

    If old < (newest - 5) Then old = old + 1: GoTo redo

End Function

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
