VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Worksheet_SelectionChange(ByVal Target As Range)
'Select Cell "A1" if game is in progress and game board is clicked on.

    'set control form focus if game board is clicked on
    If GameInProgress Then
        Wrks.Cells(1, 1).Select
        ControlForm.Hide
        ControlForm.Show vbModeless
        DoEvents
    End If

End Sub
