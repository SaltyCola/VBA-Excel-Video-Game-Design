VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frm_Controls 
   Caption         =   "Snake"
   ClientHeight    =   570
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   1890
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
    Me.Top = Application.Top + 10
    Me.Left = Application.Left + 10

End Sub

Public Sub UserForm_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)

    If KeyCode = vbKeyUp And Not DirDown And Not KeyPressed Then
        KeyPressed = True 'prevent another key press this tick
        Snakey.Head.Xvel = 0
        Snakey.Head.Yvel = -1
        DirUp = True
        DirLeft = False
        DirRight = False
        DirDown = False
    ElseIf KeyCode = vbKeyLeft And Not DirRight And Not KeyPressed Then
        KeyPressed = True 'prevent another key press this tick
        Snakey.Head.Xvel = -1
        Snakey.Head.Yvel = 0
        DirUp = False
        DirLeft = True
        DirRight = False
        DirDown = False
    ElseIf KeyCode = vbKeyRight And Not DirLeft And Not KeyPressed Then
        KeyPressed = True 'prevent another key press this tick
        Snakey.Head.Xvel = 1
        Snakey.Head.Yvel = 0
        DirUp = False
        DirLeft = False
        DirRight = True
        DirDown = False
    ElseIf KeyCode = vbKeyDown And Not DirUp And Not KeyPressed Then
        KeyPressed = True 'prevent another key press this tick
        Snakey.Head.Xvel = 0
        Snakey.Head.Yvel = 1
        DirUp = False
        DirLeft = False
        DirRight = False
        DirDown = True
    End If
    
End Sub

Public Sub UserForm_Terminate()

    'stop game counter
    GameInProg = False
    Range("G7:KT240").Interior.Color = White 'clear screen

End Sub
