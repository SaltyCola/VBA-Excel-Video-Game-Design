VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Snake_Head"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'=================================================================================='
'=================================== Properties ==================================='
'=================================================================================='

Private pXpos As Long 'horrizontal position of the head
Private pYpos As Long 'vertical position of the head
Private pXvel As Long 'horrizontal velocity of head
Private pYvel As Long 'vertical velocity of head
Private pCubeUp As DrawTool_Cube 'head sprite up
Private pCubeLeft As DrawTool_Cube 'head sprite left
Private pCubeRight As DrawTool_Cube 'head sprite right
Private pCubeDown As DrawTool_Cube 'head sprite down

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
' Head Positions
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
' Head Velocities
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
        pCubeUp.C1 = White
        pCubeUp.C2 = Black
        pCubeUp.C3 = White
        pCubeUp.C4 = Black
        pCubeUp.C5 = Black
        pCubeUp.C6 = Black
        pCubeUp.C7 = Black
        pCubeUp.C8 = Black
        pCubeUp.C9 = Black
        'Left
        pCubeLeft.C1 = White
        pCubeLeft.C2 = Black
        pCubeLeft.C3 = Black
        pCubeLeft.C4 = Black
        pCubeLeft.C5 = Black
        pCubeLeft.C6 = Black
        pCubeLeft.C7 = White
        pCubeLeft.C8 = Black
        pCubeLeft.C9 = Black
        'Right
        pCubeRight.C1 = Black
        pCubeRight.C2 = Black
        pCubeRight.C3 = White
        pCubeRight.C4 = Black
        pCubeRight.C5 = Black
        pCubeRight.C6 = Black
        pCubeRight.C7 = Black
        pCubeRight.C8 = Black
        pCubeRight.C9 = White
        'Down
        pCubeDown.C1 = Black
        pCubeDown.C2 = Black
        pCubeDown.C3 = Black
        pCubeDown.C4 = Black
        pCubeDown.C5 = Black
        pCubeDown.C6 = Black
        pCubeDown.C7 = White
        pCubeDown.C8 = Black
        pCubeDown.C9 = White

End Sub

Public Sub RedFlash()
    
    'set directions cube objects
        'Up
        pCubeUp.C1 = White
        pCubeUp.C2 = Red
        pCubeUp.C3 = White
        pCubeUp.C4 = Red
        pCubeUp.C5 = Red
        pCubeUp.C6 = Red
        pCubeUp.C7 = Red
        pCubeUp.C8 = Red
        pCubeUp.C9 = Red
        'Left
        pCubeLeft.C1 = White
        pCubeLeft.C2 = Red
        pCubeLeft.C3 = Red
        pCubeLeft.C4 = Red
        pCubeLeft.C5 = Red
        pCubeLeft.C6 = Red
        pCubeLeft.C7 = White
        pCubeLeft.C8 = Red
        pCubeLeft.C9 = Red
        'Right
        pCubeRight.C1 = Red
        pCubeRight.C2 = Red
        pCubeRight.C3 = White
        pCubeRight.C4 = Red
        pCubeRight.C5 = Red
        pCubeRight.C6 = Red
        pCubeRight.C7 = Red
        pCubeRight.C8 = Red
        pCubeRight.C9 = White
        'Down
        pCubeDown.C1 = Red
        pCubeDown.C2 = Red
        pCubeDown.C3 = Red
        pCubeDown.C4 = Red
        pCubeDown.C5 = Red
        pCubeDown.C6 = Red
        pCubeDown.C7 = White
        pCubeDown.C8 = Red
        pCubeDown.C9 = White

End Sub
