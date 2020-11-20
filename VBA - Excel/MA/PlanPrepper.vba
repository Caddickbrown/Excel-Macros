'This macro spat out multiple tabs, sorted data from a data dump tab and copied into the relevant tabs'

'Currently looking into where has gone'

'This is rough approximation of some of the functionality'

Sub AddMultiple()
'
' AddMultiple Macro
'

'
    Sheets.Add After:=ActiveSheet
    Sheets("Sheet1").Select
    Sheets("Sheet1").Name = "Fredrickson"
    Range("B2").Select
    ActiveCell.FormulaR1C1 = "Blah"
    Range("D5").Select
    Sheets("Fredrickson").Select
    ActiveWindow.SelectedSheets.Delete
    Sheets(1).Select
End Sub
