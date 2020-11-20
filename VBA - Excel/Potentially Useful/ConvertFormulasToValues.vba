'This code will convert all formulas into values
Sub ConvertToValues()
With ActiveSheet.UsedRange
.Value = .Value
End With
End Sub
