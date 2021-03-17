'This code will protect all the sheets at one go'

Sub ProtectAllSheets()

  Dim ws As Worksheet
  Dim password As String
  password = "Test123" 'replace Test123 with the password you want'

  For Each ws In Worksheets
    ws.Protect password:=password
  Next ws

End Sub
