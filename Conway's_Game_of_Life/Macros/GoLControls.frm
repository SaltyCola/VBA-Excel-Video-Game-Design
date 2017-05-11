VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} GoLControls 
   Caption         =   "Controls"
   ClientHeight    =   1800
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   3000
   OleObjectBlob   =   "GoLControls.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "GoLControls"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdStart_Click()

    'switch enabled buttons
    Me.cmdStart.Enabled = False
    Me.cmdPause.Enabled = True
    Me.cmdClear.Enabled = False
    
    'call sub
    Call Start

End Sub

Private Sub cmdPause_Click()

    'switch enabled buttons
    Me.cmdStart.Enabled = True
    Me.cmdPause.Enabled = False
    Me.cmdClear.Enabled = True
    
    'change boolean
    GameInProgress = False

End Sub

Private Sub cmdClear_Click()

    'switch enabled buttons
    Me.cmdStart.Enabled = True
    Me.cmdPause.Enabled = False
    Me.cmdClear.Enabled = True
    
    'call sub
    Call Clear

End Sub
