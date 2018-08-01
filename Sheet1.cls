VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit 'For initial error handling

'When there is a change in the sheet, look at cell D3
Private Sub Worksheet_Change(ByVal Target As Range)
If Not Intersect(Target, Range("D3")) Is Nothing Then 'Enter this conditional if the change is in D3
    Select Case Range("D3") 'Figure out what item was selected and launch the appropriate macros.
        Case "Draft": Module1.DraftChosen
        Case "90% Schematic": Module1.NinetyPercentSchematic
        Case "Final Schematic": Module1.FinalSchematic
    End Select
End If
End Sub

'All macros launched in the switch statement are found in module 1 of the file

