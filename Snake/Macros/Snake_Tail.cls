VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Snake_Tail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'=================================================================================='
'=================================== Properties ==================================='
'=================================================================================='

Private pXpos As Long 'horrizontal position of the tail
Private pYpos As Long 'vertical position of the tail
Private pXvel As Long 'horrizontal velocity of tail
Private pYvel As Long 'vertical velocity of tail
Private pCubeUp As DrawTool_Cube 'tail sprite up
Private pCubeLeft As DrawTool_Cube 'tail sprite left
Private pCubeRight As DrawTool_Cube 'tail sprite right
Private pCubeDown As DrawTool_Cube 'tail sprite down

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
' Tail Positions
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
' Tail Velocities
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

Public Property Get CubeUp() As DrawTool_Cube
    Set CubeUp = pCubeUp
End Property
Public Property Let CubeUp(Value As DrawTool_Cube)
    Set pCubeUp = Value
End Property

Public Property Get CubeLeft() As DrawTool_Cube
    Set CubeLeft = pCubeLeft
End Property
Public Property Let CubeLeft(Value As DrawTool_Cube)
    Set pCubeLeft = Value
End Property

Public Property Get CubeRight() As DrawTool_Cube
    Set CubeRight = pCubeRight
End Property
Public Property Let CubeRight(Value As DrawTool_Cube)
    Set pCubeRight = Value
End Property

Public Property Get CubeDown() As DrawTool_Cube
    Set CubeDown = pCubeDown
End Property
Public Property Let CubeDown(Value As DrawTool_Cube)
    Set pCubeDown = Value
End Property


'================================================================================'
'==================================== Methods ==================================='
'================================================================================'

Private Sub Class_Initialize()
    
    'initialize directional cube objects
    Set pCubeUp = New DrawTool_Cube
    Set pCubeLeft = New DrawTool_Cube
    Set pCubeRight = New DrawTool_Cube
    Set pCubeDown = New DrawTool_Cube
    
End Sub

Public Sub SetCubes()
    
    'set directions cube objects
        'Up
        pCubeUp.C1 = Black
        pCubeUp.C2 = Black
        pCubeUp.C3 = Black
        pCubeUp.C4 = Black
        pCubeUp.C5 = Black
        pCubeUp.C6 = Black
        pCubeUp.C7 = Black
        pCubeUp.C8 = White
        pCubeUp.C9 = Black
        'Left
        pCubeLeft.C1 = Black
        pCubeLeft.C2 = Black
        pCubeLeft.C3 = Black
        pCubeLeft.C4 = Black
        pCubeLeft.C5 = Black
        pCubeLeft.C6 = White
        pCubeLeft.C7 = Black
        pCubeLeft.C8 = Black
        pCubeLeft.C9 = Black
        'Right
        pCubeRight.C1 = Black
        pCubeRight.C2 = Black
        pCubeRight.C3 = Black
        pCubeRight.C4 = White
        pCubeRight.C5 = Black
        pCubeRight.C6 = Black
        pCubeRight.C7 = Black
        pCubeRight.C8 = Black
        pCubeRight.C9 = Black
        'Down
        pCubeDown.C1 = Black
        pCubeDown.C2 = White
        pCubeDown.C3 = Black
        pCubeDown.C4 = Black
        pCubeDown.C5 = Black
        pCubeDown.C6 = Black
        pCubeDown.C7 = Black
        pCubeDown.C8 = Black
        pCubeDown.C9 = Black

End Sub

Public Sub RedFlash()
    
    'set directions cube objects
        'Up
        pCubeUp.C1 = Red
        pCubeUp.C2 = Red
        pCubeUp.C3 = Red
        pCubeUp.C4 = Red
        pCubeUp.C5 = Red
        pCubeUp.C6 = Red
        pCubeUp.C7 = Red
        pCubeUp.C8 = White
        pCubeUp.C9 = Red
        'Left
        pCubeLeft.C1 = Red
        pCubeLeft.C2 = Red
        pCubeLeft.C3 = Red
        pCubeLeft.C4 = Red
        pCubeLeft.C5 = Red
        pCubeLeft.C6 = White
        pCubeLeft.C7 = Red
        pCubeLeft.C8 = Red
        pCubeLeft.C9 = Red
        'Right
        pCubeRight.C1 = Red
        pCubeRight.C2 = Red
        pCubeRight.C3 = Red
        pCubeRight.C4 = White
        pCubeRight.C5 = Red
        pCubeRight.C6 = Red
        pCubeRight.C7 = Red
        pCubeRight.C8 = Red
        pCubeRight.C9 = Red
        'Down
        pCubeDown.C1 = Red
        pCubeDown.C2 = White
        pCubeDown.C3 = Red
        pCubeDown.C4 = Red
        pCubeDown.C5 = Red
        pCubeDown.C6 = Red
        pCubeDown.C7 = Red
        pCubeDown.C8 = Red
        pCubeDown.C9 = Red

End Sub
