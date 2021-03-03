'This Macro gives the basis for a form that can be filled out as a template

Private Sub CommandButton2_Click()
End Sub

Private Sub cancelBut_Click()
    stInfo.Hide
End Sub

Private Sub Label2_Click()
End Sub

Private Sub OKbut_Click()
    Dim studentName As Range
    Set studentName = ActiveDocument.Bookmarks("sName").Range
    studentName.Text = Me.TextBox1.Value
    Dim schoolName As Range
    Set schoolName = ActiveDocument.Bookmarks("sSchool").Range
    schoolName.Text = Me.TextBox3.Value
    Dim paperTitle As Range
    Dim p2Title As Range
    Dim hTitle As Range
    Dim h2Title As Range
    Set paperTitle = ActiveDocument.Bookmarks("pTitle").Range
    Set h2Title = ActiveDocument.Bookmarks("h2Title").Range
    Set hTitle = ActiveDocument.Bookmarks("hTitle").Range
    Set p2Title = ActiveDocument.Bookmarks("p2Title").Range
    paperTitle.Text = Me.TextBox2.Value
    p2Title.Text = Me.TextBox2.Value
    h2Title.Text = Me.TextBox2.Value
    hTitle.Text = Me.TextBox2.Value
    hTitle.Font.AllCaps = True
    h2Title.Font.AllCaps = True
    Me.Repaint
    Dim strDocName As String
    Dim intPos As Integer

    'Find position of extension in file name
    strDocName = ""
    intPos = InStrRev(strDocName, ".")

    If intPos = 0 Then
    ' If the document has not yet been saved - Ask the user to provide a file name
        strDocName = InputBox("Please enter the name " & _
            "of your document.")
    Else
        ' Strip off extension and add ".txt" extension
        strDocName = Left(strDocName, intPos - 1)
        strDocName = strDocName & ".docx"
    End If
    ' Save file with new extension
    ActiveDocument.SaveAs2 FileName:=strDocName, _
        FileFormat:=wdFormatDocumentDefault
    stInfo.Hide
    infoForm.Show

End Sub
