VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Model_Pellet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'=================================================================================='
'=================================== Properties ==================================='
'=================================================================================='

Private pXpos As Long 'horrizontal position of the pellet
Private pYpos As Long 'vertical position of the pellet
Private pCubeAny As DrawTool_Cube 'pellet sprite

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
' Pellet Positions
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
' Cube Sprite
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
        pCubeAny.C2 = Black
        pCubeAny.C3 = White
        pCubeAny.C4 = Black
        pCubeAny.C5 = Black
        pCubeAny.C6 = Black
        pCubeAny.C7 = White
        pCubeAny.C8 = Black
        pCubeAny.C9 = White

End Sub

Public Sub Draw()

    Dim xUpperBound As Integer 'For random number generator
    Dim xLowerBound As Integer 'For random number generator
    Dim yUpperBound As Integer 'For random number generator
    Dim yLowerBound As Integer 'For random number generator
    
    'initialize bounds
    xLowerBound = 0
    xUpperBound = (300 / 3) - 3
    yLowerBound = 0
    yUpperBound = (234 / 3) - 3
    
    'randomly choose pellet position
        Randomize 'generate seed value for Rnd
    Me.Xpos = 3 * Int((xUpperBound - xLowerBound + 1) * Rnd + xLowerBound)
        Randomize 'generate seed value for Rnd
    Me.Ypos = 3 * Int((yUpperBound - yLowerBound + 1) * Rnd + yLowerBound)
    
    'draw pellet
    Me.CubeAny.DrawAt (yMin + Me.Ypos), (xMin + Me.Xpos)

End Sub
