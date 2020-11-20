'Seems to copy a sheet? Not really sure'

Sub CTPTRIALING()

    Worksheets(1).Activate
    Worksheets(1).Copy Before:=Sheets(2)
    Dim rs As Worksheet

    For Each rs In Sheets
    rs.Name = rs.Range("A1")
    Next rs

End Sub
