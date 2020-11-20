

Sub NotesMacro()
'
' NotesMacro Macro
'

    Application.ScreenUpdating = False


    Worksheets("Gantt").Select
    Range("A1").Activate

    Worksheets("Notes").ShowDataForm
    Application.ScreenUpdating = True

End Sub
