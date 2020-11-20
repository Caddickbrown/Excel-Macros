'This code will protect all sheets in the workbook
Sub ProtectAllSheets()
Dim ws As Worksheet
For Each ws In Worksheets
ws.Protect
Next ws
End Sub
