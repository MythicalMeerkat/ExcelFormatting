Attribute VB_Name = "Module1"
Sub DraftChosen()

'Resets all columns temporarily, without this columns would continue to be hidden when we want them to be shown
Columns("E:G").EntireColumn.Hidden = False


'Hide appropriate columns
Columns("F:G").EntireColumn.Hidden = True

End Sub
Sub NinetyPercentSchematic()


'Resets all columns temporarily, without this columns would continue to be hidden when we want them to be shown
Columns("E:G").EntireColumn.Hidden = False


'Hide appropriate columns
Columns("E").EntireColumn.Hidden = True
Columns("G").EntireColumn.Hidden = True

End Sub

Sub FinalSchematic()

'Resets all columns temporarily, without this columns would continue to be hidden when we want them to be shown
Columns("E:G").EntireColumn.Hidden = False


'Hide appropriate columns
Columns("E:F").EntireColumn.Hidden = True

End Sub




