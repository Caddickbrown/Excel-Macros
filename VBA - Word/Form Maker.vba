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



' Below is in progress work


'This Macro gives the basis for a form that can be filled out as a template

Private Sub CommandButton2_Click()
End Sub

Private Sub animalBox1_Change()
    ComboBox1.AddItem ("Human")
    ComboBox1.AddItem ("Canine")
    ComboBox1.AddItem ("Fish")
    ComboBox1.AddItem ("Feline")
    ComboBox1.AddItem ("Bovine")
End Sub

Private Sub bioBox_Change()
With ProjectDetails.bioBox
        .AddItem ("blood")
        .AddItem ("plasma")
        .AddItem ("serum")
        .AddItem ("urine")
        .AddItem ("nasal")
    End With
End Sub

Sub jImmy()
    stInfo.Show
End Sub

Private Sub cancelBut_Click()
    ProjectDetails.Hide
End Sub

Private Sub ComboBox1_Change()

End Sub

Private Sub interpretBox1_Change()
    ComboBox1.AddItem ("visually")
    ComboBox1.AddItem ("reader")
End Sub


Private Sub OKbut_Click()

    Dim companyName As Range
    Set companyName = ActiveDocument.Bookmarks("companyName").Range
    companyName.Text = Me.TextBox1.Value

    Dim biomarkerName As Range
    Set biomarkerName = ActiveDocument.Bookmarks("biomarkerName").Range
    biomarkerName.Text = Me.TextBox3.Value

    Dim sampleType As Range
    Set sampleType = ActiveDocument.Bookmarks("sampleType").Range
    sampleType.Text = Me.bioBox.Value

    Dim conditionDisease As Range
    Set conditionDisease = ActiveDocument.Bookmarks("conditionDisease").Range
    conditionDisease.Text = Me.TextBox5.Value

    Dim animalName As Range
    Set animalName = ActiveDocument.Bookmarks("animalName").Range
    animalName.Text = Me.animalBox1.Value

    Dim interpretingMethod As Range
    Set interpretingMethod = ActiveDocument.Bookmarks("interpretingMethod").Range
    interpretingMethod.Text = Me.interpretBox1.Value

    Me.Repaint
    ProjectDetails.Hide

End Sub

Private Sub TextBox3_Change()

End Sub
