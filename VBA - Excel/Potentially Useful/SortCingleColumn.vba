'This will sort the data by a single column'
Sub SortDataHeader()
Range("DataRange").Sort Key1:=Range("A1"), Order1:=xlAscending, Header:=xlYes
End Sub
