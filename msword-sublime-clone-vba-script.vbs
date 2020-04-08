
Sub DimKeyword()

	Dim word As String
	Dim color As Long


	word = "Dim"
	color = RGB(0, 255, 0)

	Application.DisplayAlerts = False

	Selection.WholeStory
	Selection.Find.ClearFormatting
	Selection.Find.Replacement.ClearFormatting
	Selection.Find.Replacement.Font.Color = color

	With Selection.Find

		.text = word
		.Replacement.Text = ""
		.Forward = True
		.Wrap = wdFindAsk
		.Format = True
		.MatchCase = True
		.MatchWholeWord = True

	End With

	Selection.Find.Execute Replace:=wdReplaceAll

End Sub
==============================================================
Sub StringDataType

	Dim word As String
	Dim color As Long

	' escaping the double quote by using ASCII 34
	word = "[" & Chr(34) & "|“]*[" & Chr(34) & "|”]"
	color = RGB(0, 255, 0)

	Application.DisplayAlerts = False

	Selection.WholeStory
	Selection.Find.ClearFormatting
	Selection.Find.Replacement.ClearFormatting
	Selection.Find.Replacement.Font.Color = color

	With Selection.Find

		.text = word
		.Replacement.Text = ""
		.Forward = True
		.Wrap = wdFindAsk
		.Format = True
		.MatchCase = True
		.MatchWholeWord = True

	End With

	Selection.Find.Execute Replace:=wdReplaceAll

End Sub