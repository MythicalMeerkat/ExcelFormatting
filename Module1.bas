Attribute VB_Name = "Module1"
Sub DraftChosen()

'MsgBox "Draft Chosen" :: FOR TESTING PURPOSES ONLY


'Resets all columns temporarily, without this columns would continue to be hidden when we want them to be shown
Columns("E:G").EntireColumn.Hidden = False


'Hide appropriate columns
Columns("F:G").EntireColumn.Hidden = True

End Sub

Sub NinetyPercentSchematic()

'MsgBox "90% Schematic Chosen":: FOR TESTING PUEPOSES ONLY


'Resets all columns temporarily, without this columns would continue to be hidden when we want them to be shown
Columns("E:G").EntireColumn.Hidden = False


'Hide appropriate columns
Columns("E").EntireColumn.Hidden = True
Columns("G").EntireColumn.Hidden = True

End Sub

Sub FinalSchematic()

'MsgBox "Final Schematic Chosen" :: FOR TESTING PURPOSES ONLY


'Resets all columns temporarily, without this columns would continue to be hidden when we want them to be shown
Columns("E:G").EntireColumn.Hidden = False


'Hide appropriate columns
Columns("E:F").EntireColumn.Hidden = True

End Sub


