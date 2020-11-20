'Used to clear a template for data entry'

Sub Clear_CTP()
'
' Clear_CTP Macro
'
'
    Range("B16:C16").Select
    Range("C16").Activate
    Selection.ClearContents
    Range("H16:I16").Select
    Selection.ClearContents
    Range("P16").Select
    Selection.ClearContents
    Range("J16").Select
    Selection.ClearContents
    Range("N16").Select
    Selection.ClearContents
    Range("M16").Select
    Selection.ClearContents
    Range("B16").Select
End Sub
