VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Snake_Eraser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

'=================================================================================='
'=================================== Properties ==================================='
'=================================================================================='

Private pXpos As Long 'horrizontal position of the eraser
Private pYpos As Long 'vertical position of the eraser
Private pXvel As Long 'horrizontal velocity of eraser
Private pYvel As Long 'vertical velocity of eraser
Private pCubeAny As DrawTool_Cube 'eraser sprite

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
' Eraser Positions
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
' Eraser Velocities
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
        'Any
        pCubeAny.C1 = White
        pCubeAny.C2 = White
        pCubeAny.C3 = White
        pCubeAny.C4 = White
        pCubeAny.C5 = White
        pCubeAny.C6 = White
        pCubeAny.C7 = White
        pCubeAny.C8 = White
        pCubeAny.C9 = White

End Sub

