'This Macro gives the basis for a form that can be filled out as a template - this will need a "UserForm1" creating and any options will need customising.

Sub FormMaker()
    UserForm1.Show
End Sub

Private Sub UserForm_Initialize()
    ComboBox1.AddItem ("IgG")
    ComboBox1.AddItem ("IgM")
    ComboBox2.AddItem ("blood")
    ComboBox2.AddItem ("plasma")
    ComboBox2.AddItem ("serum")
    ComboBox2.AddItem ("urine")
    ComboBox2.AddItem ("nasal")
    ComboBox3.AddItem ("Human")
    ComboBox3.AddItem ("Canine")
    ComboBox3.AddItem ("Fish")
    ComboBox3.AddItem ("Feline")
    ComboBox3.AddItem ("Bovine")
    ComboBox4.AddItem ("visually")
    ComboBox4.AddItem ("reader")
End Sub

Private Sub CommandButton2_Click()
    UserForm1.Hide
End Sub

Private Sub CommandButton1_Click()

    CommandBars("Navigation").Visible = False
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "CompanyName"
        .Replacement.Text = Me.TextBox1.Value
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll

    CommandBars("Navigation").Visible = False
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
      .Text = "BiomarkerName"
      .Replacement.Text = Me.ComboBox1.Value
      .Forward = True
      .Wrap = wdFindContinue
      .Format = False
      .MatchCase = False
      .MatchWholeWord = False
      .MatchWildcards = False
      .MatchSoundsLike = False
      .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll

    CommandBars("Navigation").Visible = False
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
      .Text = "SampleType"
      .Replacement.Text = Me.ComboBox2.Value
      .Forward = True
      .Wrap = wdFindContinue
      .Format = False
      .MatchCase = False
      .MatchWholeWord = False
      .MatchWildcards = False
      .MatchSoundsLike = False
      .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll

    CommandBars("Navigation").Visible = False
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
      .Text = "ConditionDisease"
      .Replacement.Text = Me.TextBox2.Value
      .Forward = True
      .Wrap = wdFindContinue
      .Format = False
      .MatchCase = False
      .MatchWholeWord = False
      .MatchWildcards = False
      .MatchSoundsLike = False
      .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll

    CommandBars("Navigation").Visible = False
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
      .Text = "AnimalName"
      .Replacement.Text = Me.ComboBox3.Value
      .Forward = True
      .Wrap = wdFindContinue
      .Format = False
      .MatchCase = False
      .MatchWholeWord = False
      .MatchWildcards = False
      .MatchSoundsLike = False
      .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll


    CommandBars("Navigation").Visible = False
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
      .Text = "InterpretingMethod"
      .Replacement.Text = Me.ComboBox4.Value
      .Forward = True
      .Wrap = wdFindContinue
      .Format = False
      .MatchCase = False
      .MatchWholeWord = False
      .MatchWildcards = False
      .MatchSoundsLike = False
      .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll

    UserForm1.Hide

End Sub
