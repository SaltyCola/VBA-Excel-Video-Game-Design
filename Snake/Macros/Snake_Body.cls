VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Snake_Body"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'=================================================================================='
'=================================== Properties ==================================='
'=================================================================================='

Private pXpos As Long 'horrizontal position of the body piece
Private pYpos As Long 'vertical position of the body piece
Private pXvel As Long 'horrizontal velocity of body piece
Private pYvel As Long 'vertical velocity of body piece
Private pCubeAny As DrawTool_Cube 'body sprite

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
' Body Positions
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Property Get Xpos() As Long
    Xpos = pXpos
End Property
Public Property Let Xpos(Value As Long)
    pXpos = Value
End Property

Public Property Get Ypos() As Long
    Ypos = pYpos
End Property
Public Property Let Ypos(Value As Long)
    pYpos = Value
End Property

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
' Body Velocities
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Property Get Xvel() As Long
    Xvel = pXvel
End Property
Public Property Let Xvel(Value As Long)
    pXvel = Value
End Property

Public Property Get Yvel() As Long
    Yvel = pYvel
End Property
Public Property Let Yvel(Value As Long)
    pYvel = Value
End Property

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
' Directional Cube Sprites
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Property Get CubeAny() As DrawTool_Cube
    Set CubeAny = pCubeAny
End Property
Public Property Let CubeAny(Value As DrawTool_Cube)
    Set pCubeAny = Value
End Property


'================================================================================'
'==================================== Methods ==================================='
'================================================================================'

Private Sub Class_Initialize()
    
    'initialize directional cube objects
    Set pCubeAny = New DrawTool_Cube
    
End Sub

Public Sub SetCubes()
    
    'set directions cube objects
        'Any Direction
        pCubeAny.C1 = Black
        pCubeAny.C2 = Black
        pCubeAny.C3 = Black
        pCubeAny.C4 = Black
        pCubeAny.C5 = Black
        pCubeAny.C6 = Black
        pCubeAny.C7 = Black
        pCubeAny.C8 = Black
        pCubeAny.C9 = Black

End Sub

Public Sub RedFlash()
    
    'set directions cube objects
        'Any Direction
        pCubeAny.C1 = Red
        pCubeAny.C2 = Red
        pCubeAny.C3 = Red
        pCubeAny.C4 = Red
        pCubeAny.C5 = Red
        pCubeAny.C6 = Red
        pCubeAny.C7 = Red
        pCubeAny.C8 = Red
        pCubeAny.C9 = Red

End Sub
