VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frm_Controls 
   Caption         =   "Ms. Pacman"
   ClientHeight    =   480
   ClientLeft      =   48
   ClientTop       =   372
   ClientWidth     =   1884
   OleObjectBlob   =   "frm_Controls.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frm_Controls"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Sub UserForm_Activate()

    'Make Control Board Form appear in top left corner of screen
    Me.StartUpPosition = 0
    Me.top = Application.top + 5
    Me.left = Application.left + 5
    Me.Width = 0
    Me.Height = 0

End Sub

Public Sub UserForm_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
'allows player to hold a key to change MsP's direction

    Dim i As Integer 'integer iterator
    
    'log key press
    For i = 1 To 8
        If KeyCode = KeyState(i, 1) Then
            KeyState(i, 2) = True
        End If
    Next i
    
End Sub

Public Sub UserForm_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
'Ensures MsP will only change direction while a directional key is held down

    Dim i As Integer 'integer iterator
    
    'log key press
    For i = 1 To 8
        If KeyCode = KeyState(i, 1) Then
            KeyState(i, 2) = False
        End If
    Next i
    
End Sub

Public Sub UserForm_Terminate()

    'call game end
    Call GameEnd

End Sub
