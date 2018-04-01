Attribute VB_Name = "NewMacros"
Sub change()

    Selection.WholeStory
    Selection.Font.Name = "Times New Roman"

    For i = 61632 To 61695
        a1 = i
        a = Trim("^u") & Trim(Str(a1))
        b1 = i - 60592
        Selection.Find.ClearFormatting
        Selection.Find.Replacement.ClearFormatting
        With Selection.Find
            .Text = a
            .Replacement.Text = ChrW(b1)
            .Forward = True
            .Wrap = wdFindContinue
            .MatchCase = True
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
    Next i


    For i = 61472 To 61533
        a1 = i
        a = Trim("^u") & Trim(Str(a1))
        b1 = i - 61440
        Selection.Find.ClearFormatting
        Selection.Find.Replacement.ClearFormatting
        With Selection.Find
            .Text = a
            .Replacement.Text = ChrW(b1)
            .Forward = True
            .Wrap = wdFindContinue
            .MatchCase = True
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
    Next i

    For i = 61535 To 61627
        a1 = i
        a = Trim("^u") & Trim(Str(a1))
        b1 = i - 61440
        Selection.Find.ClearFormatting
        Selection.Find.Replacement.ClearFormatting
        With Selection.Find
            .Text = a
            .Replacement.Text = ChrW(b1)
            .Forward = True
            .Wrap = wdFindContinue
            .MatchCase = True
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
    Next i

    Selection.WholeStory
    Selection.LanguageID = wdRussian
    Application.CheckLanguage = True
End Sub
