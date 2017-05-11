VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "AI_Ghosts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'=================================================================================='
'=================================== Properties ==================================='
'=================================================================================='

'AI Movement Type
Private pMoveType As Integer 'Tells Ghost what movement type it currently is in (0=Cage ; 1=Chase ; 2=Scatter ; 3=Scared ; 4=Eyes)
Private pCageMoveType As Integer 'Tells Ghost how to move in cage (0=not in cage ; 1=enter ; 2=wiggle ; 3=exit)
Private pCageMoveCounter As Integer 'Counts which anim frame cage move is on (there is no track in cage to prevent MsP from entering)

'Timers
Private pTimerCage As Integer 'Tells Ghost how long to wait until out of cage (10 frames ~ 1 second)

'Object Coordinates
Private pXobject As Integer 'object's x or column coordinate
Private pYobject As Integer 'object's y or row coordinate

'Target Coordinates
Private pXtarget As Integer 'target's x or column coordinate
Private pYtarget As Integer 'target's y or row coordinate

'Sight Line
Private pSLup As Double 'Distance from up turn to target
Private pSLleft As Double 'Distance from left turn to target
Private pSLright As Double 'Distance from right turn to target
Private pSLdown As Double 'Distance from down turn to target

'Directions [1=Up ; 2=Left ; 3=Right ; 4=Down]
Private pDirLast As Integer 'Last direction taken (helps calculate which direction not to take on next turn)
Private pDirPrevent As Integer 'Prevents AI from not choosing this direction on next turn

'''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
' AI Movement Type
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Property Get MoveType() As Integer
    MoveType = pMoveType
End Property
Public Property Let MoveType(Value As Integer)
    pMoveType = Value
End Property

'''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
' AI Ghost Cage Movement Type
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Property Get CageMoveType() As Integer
    CageMoveType = pCageMoveType
End Property
Public Property Let CageMoveType(Value As Integer)
    pCageMoveType = Value
End Property

'''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
' AI Ghost Cage Movement Counter
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Property Get CageMoveCounter() As Integer
    CageMoveCounter = pCageMoveCounter
End Property
Public Property Let CageMoveCounter(Value As Integer)
    pCageMoveCounter = Value
End Property

'''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
' Timer Cage Duration
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Property Get TimerCage() As Integer
    TimerCage = pTimerCage
End Property
Public Property Let TimerCage(Value As Integer)
    pTimerCage = Value
End Property

'''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
' Circular Motion Type
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Property Get CircularMotion() As Integer
    CircularMotion = pCircularMotion
End Property
Public Property Let CircularMotion(Value As Integer)
    pCircularMotion = Value
End Property

'''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
' Object's X or Col coord
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Property Get Xobject() As Integer
    Xobject = pXobject
End Property
Public Property Let Xobject(Value As Integer)
    pXobject = Value
End Property

'''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
' Object's Y or Row coord
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Property Get Yobject() As Integer
    Yobject = pYobject
End Property
Public Property Let Yobject(Value As Integer)
    pYobject = Value
End Property

'''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
' Target's X or Col coord
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Property Get Xtarget() As Integer
    Xtarget = pXtarget
End Property
Public Property Let Xtarget(Value As Integer)
    pXtarget = Value
End Property

'''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
' Target's Y or Row coord
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Property Get Ytarget() As Integer
    Ytarget = pYtarget
End Property
Public Property Let Ytarget(Value As Integer)
    pYtarget = Value
End Property

'''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
' Sight Line Up Distance
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Property Get SLup() As Double
    SLup = pSLup
End Property
Public Property Let SLup(Value As Double)
    pSLup = Value
End Property

'''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
' Sight Line Left Distance
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Property Get SLleft() As Double
    SLleft = pSLleft
End Property
Public Property Let SLleft(Value As Double)
    pSLleft = Value
End Property

'''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
' Sight Line Right Distance
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Property Get SLright() As Double
    SLright = pSLright
End Property
Public Property Let SLright(Value As Double)
    pSLright = Value
End Property

'''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
' Sight Line Down Distance
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Property Get SLdown() As Double
    SLdown = pSLdown
End Property
Public Property Let SLdown(Value As Double)
    pSLdown = Value
End Property

'''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
' Last Direction Change Taken
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Property Get DirLast() As Integer
    DirLast = pDirLast
End Property
Public Property Let DirLast(Value As Integer)
    pDirLast = Value
End Property

'''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
' Prevent Taking This Direction
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Property Get DirPrevent() As Integer
    DirPrevent = pDirPrevent
End Property
Public Property Let DirPrevent(Value As Integer)
    pDirPrevent = Value
End Property


'================================================================================='
'================================ Public Methods ================================='
'================================================================================='

Public Sub TargetChange(ByVal AI_Object As Sprite_Ghosts)
'Change coordinates of the target

    'Blinky
    If AI_Object.GhostName = "Blinky" Then
        'Chase
        If pMoveType = 1 Then
            pXtarget = MsP.Xcol
            pYtarget = MsP.Yrow
        'Scatter
        ElseIf pMoveType = 2 Then 'top-right corner
            pXtarget = 237
            pYtarget = 13
        'Scared
        ElseIf pMoveType = 3 Then 'move randomly
            pXtarget = MsP.Xcol
            pYtarget = MsP.Yrow
        'Eyes
        ElseIf pMoveType = 4 Then 'head for ghost cage entrance
            pXtarget = 121
            pYtarget = 109
        End If
    
    'Pinky
    ElseIf AI_Object.GhostName = "Pinky" Then
        'Chase
        If pMoveType = 1 Then
            'MsP moving up
            If MsP.nDir = 1 Then
                pXtarget = MsP.Xcol
                pYtarget = MsP.Yrow - 32 '4 cubes in front
            'MsP moving left
            ElseIf MsP.nDir = 2 Then
                pXtarget = MsP.Xcol - 32 '4 cubes in front
                pYtarget = MsP.Yrow
            'MsP moving right
            ElseIf MsP.nDir = 3 Then
                pXtarget = MsP.Xcol + 32 '4 cubes in front
                pYtarget = MsP.Yrow
            'MsP moving down
            ElseIf MsP.nDir = 4 Then
                pXtarget = MsP.Xcol
                pYtarget = MsP.Yrow + 32 '4 cubes in front
            End If
        'Scatter
        ElseIf pMoveType = 2 Then 'top-left corner
            pXtarget = 5
            pYtarget = 13
        'Scared
        ElseIf pMoveType = 3 Then 'move randomly
            pXtarget = MsP.Xcol
            pYtarget = MsP.Yrow
        'Eyes
        ElseIf pMoveType = 4 Then 'head for ghost cage entrance
            pXtarget = 121
            pYtarget = 109
        End If
    
    'Inky
    ElseIf AI_Object.GhostName = "Inky" Then
        'Chase
        If pMoveType = 1 Then
            'MsP moving up
            If MsP.nDir = 1 Then
                pXtarget = MsP.Xcol + (2 * (MsP.Xcol - Blinky.Xcol))
                pYtarget = MsP.Yrow - 16 + (2 * (MsP.Yrow - Blinky.Yrow))
            'MsP moving left
            ElseIf MsP.nDir = 2 Then
                pXtarget = MsP.Xcol - 16 + (2 * (MsP.Xcol - Blinky.Xcol))
                pYtarget = MsP.Yrow + (2 * (MsP.Yrow - Blinky.Yrow))
            'MsP moving right
            ElseIf MsP.nDir = 3 Then
                pXtarget = MsP.Xcol + 16 + (2 * (MsP.Xcol - Blinky.Xcol))
                pYtarget = MsP.Yrow + (2 * (MsP.Yrow - Blinky.Yrow))
            'MsP moving down
            ElseIf MsP.nDir = 4 Then
                pXtarget = MsP.Xcol + (2 * (MsP.Xcol - Blinky.Xcol))
                pYtarget = MsP.Yrow + 16 + (2 * (MsP.Yrow - Blinky.Yrow))
            End If
        'Scatter
        ElseIf pMoveType = 2 Then 'bottom-right corner
            pXtarget = 237
            pYtarget = 269
        'Scared
        ElseIf pMoveType = 3 Then 'move randomly
            pXtarget = MsP.Xcol
            pYtarget = MsP.Yrow
        'Eyes
        ElseIf pMoveType = 4 Then 'head for ghost cage entrance
            pXtarget = 121
            pYtarget = 109
        End If
    
    'Sue
    ElseIf AI_Object.GhostName = "Sue" Then
        'Chase
        If pMoveType = 1 Then
            'Greater than 8 cubes away
            If (Abs(pXobject - MsP.Xcol) > 64) Or (Abs(pYobject - MsP.Yrow) > 64) Then
                pXtarget = MsP.Xcol
                pYtarget = MsP.Yrow
            'Too Close to MsP
            Else
                pXtarget = 5
                pYtarget = 269
            End If
        'Scatter
        ElseIf pMoveType = 2 Then 'bottom-left corner
            pXtarget = 5
            pYtarget = 269
        'Scared
        ElseIf pMoveType = 3 Then 'move randomly
            pXtarget = MsP.Xcol
            pYtarget = MsP.Yrow
        'Eyes
        ElseIf pMoveType = 4 Then 'head for ghost cage entrance
            pXtarget = 121
            pYtarget = 109
        End If
    End If

End Sub

Public Sub CalculateSightLines(ByRef AI_Object As Sprite_Ghosts)
'Calculate coordinates of food object and sight lines

    Dim deltaXup As Double 'assist in calculating sight line distances
    Dim deltaYup As Double
    Dim deltaXleft As Double
    Dim deltaYleft As Double
    Dim deltaXright As Double
    Dim deltaYright As Double
    Dim deltaXdown As Double
    Dim deltaYdown As Double
    
    'AI Object Coordinates
        pXobject = AI_Object.Xcol
        pYobject = AI_Object.Yrow
    
    'Deltas
        'Up
        deltaXup = pXtarget - (pXobject)
        deltaYup = pYtarget - (pYobject - AI_Object.Speed)
        'Left
        deltaXleft = pXtarget - (pXobject - AI_Object.Speed)
        deltaYleft = pYtarget - (pYobject)
        'Right
        deltaXright = pXtarget - (pXobject + AI_Object.Speed)
        deltaYright = pYtarget - (pYobject)
        'Down
        deltaXdown = pXtarget - (pXobject)
        deltaYdown = pYtarget - (pYobject + AI_Object.Speed)
    
    'Sight Line Distances
        pSLup = Sqr(deltaXup ^ 2 + deltaYup ^ 2)
        pSLleft = Sqr(deltaXleft ^ 2 + deltaYleft ^ 2)
        pSLright = Sqr(deltaXright ^ 2 + deltaYright ^ 2)
        pSLdown = Sqr(deltaXdown ^ 2 + deltaYdown ^ 2)

End Sub

Public Sub DetermineDirection(ByRef AI_Object As Sprite_Ghosts, ByVal bUp As Boolean, ByVal bLeft As Boolean, ByVal bRight As Boolean, ByVal bDown As Boolean)
'Determines which direction the food object should take when it reaches an intersection.

    Dim intRnd As Integer 'random choice for two correct decisions (will never be more than 2)
    Dim intDir As Integer 'Direction of next turn [1=up ; 2=left ; 3=right ; 4=down]
    Dim b1 As Boolean 'up is a correct choice
    Dim b2 As Boolean 'left is a correct choice
    Dim b3 As Boolean 'right is a correct choice
    Dim b4 As Boolean 'down is a correct choice
    Dim sl_chosen As Double 'chosen sightline out of all 4 sightlines
    Dim cc As Integer 'Number of correct choices
    
    'initialize correct choice booleans
    b1 = False
    b2 = False
    b3 = False
    b4 = False
    cc = 0
    
    'initialize intDir
    intDir = 0
    
    'Calculate shortest distance
    If pMoveType <> 3 Then
        'Initialize to value higher than possible
        sl_chosen = 5000
        'Up
        If pSLup < sl_chosen And bUp And pDirPrevent <> 1 Then: sl_chosen = pSLup
        'Left
        If pSLleft < sl_chosen And bLeft And pDirPrevent <> 2 Then: sl_chosen = pSLleft
        'Right
        If pSLright < sl_chosen And bRight And pDirPrevent <> 3 Then: sl_chosen = pSLright
        'Down
        If pSLdown < sl_chosen And bDown And pDirPrevent <> 4 Then: sl_chosen = pSLdown
    
    'Calculate random sight line in Scared State
    ElseIf pMoveType = 3 Then
        'Random Choice
            Randomize
            intRnd = Int((100 - 1 + 1) * Rnd + 1) 'int((upperbound-lowerbound+1)*Rnd+lowerbound)
        'Up First
        If intRnd <= 25 Then
            If pDirPrevent <> 1 And bUp Then 'Up
                sl_chosen = pSLup
            ElseIf pDirPrevent <> 2 And bLeft Then 'Left
                sl_chosen = pSLleft
            ElseIf pDirPrevent <> 3 And bRight Then 'Right
                sl_chosen = pSLright
            ElseIf pDirPrevent <> 4 And bDown Then 'Down
                sl_chosen = pSLdown
            End If
        'Left First
        ElseIf intRnd > 25 And intRnd <= 50 Then
            If pDirPrevent <> 2 And bLeft Then 'Left
                sl_chosen = pSLleft
            ElseIf pDirPrevent <> 3 And bRight Then 'Right
                sl_chosen = pSLright
            ElseIf pDirPrevent <> 4 And bDown Then 'Down
                sl_chosen = pSLdown
            ElseIf pDirPrevent <> 1 And bUp Then 'Up
                sl_chosen = pSLup
            End If
        'Right First
        ElseIf intRnd > 50 And intRnd <= 75 Then
            If pDirPrevent <> 3 And bRight Then 'Right
                sl_chosen = pSLright
            ElseIf pDirPrevent <> 4 And bDown Then 'Down
                sl_chosen = pSLdown
            ElseIf pDirPrevent <> 1 And bUp Then 'Up
                sl_chosen = pSLup
            ElseIf pDirPrevent <> 2 And bLeft Then 'Left
                sl_chosen = pSLleft
            End If
        'Down First
        ElseIf intRnd > 75 Then
            If pDirPrevent <> 4 And bDown Then 'Down
                sl_chosen = pSLdown
            ElseIf pDirPrevent <> 1 And bUp Then 'Up
                sl_chosen = pSLup
            ElseIf pDirPrevent <> 2 And bLeft Then 'Left
                sl_chosen = pSLleft
            ElseIf pDirPrevent <> 3 And bRight Then 'Right
                sl_chosen = pSLright
            End If
        End If
    End If
    
    'Get Correct Choices
        'Up
        If pSLup = sl_chosen And bUp And pDirPrevent <> 1 Then
            b1 = True
            intDir = 1
            cc = cc + 1
        End If
        'Left
        If pSLleft = sl_chosen And bLeft And pDirPrevent <> 2 Then
            b2 = True
            intDir = 2
            cc = cc + 1
        End If
        'Right
        If pSLright = sl_chosen And bRight And pDirPrevent <> 3 Then
            b3 = True
            intDir = 3
            cc = cc + 1
        End If
        'Down
        If pSLdown = sl_chosen And bDown And pDirPrevent <> 4 Then
            b4 = True
            intDir = 4
            cc = cc + 1
        End If
    
    '2 Correct Choices
        If cc = 2 Then '(should be at most 2 correct choices)
            'Random Choice
                Randomize
                intRnd = Int((100 - 1 + 1) * Rnd + 1) 'int((upperbound-lowerbound+1)*Rnd+lowerbound)
            'Choice 1
            If intRnd <= 50 Then
                'Up
                If b1 Then
                    intDir = 1
                'Left
                ElseIf b2 Then
                    intDir = 2
                'Right
                ElseIf b3 Then
                    intDir = 3
                'Down
                ElseIf b4 Then
                    intDir = 4
                End If
            'Choice 2
            ElseIf intRnd > 50 Then
                'Down
                If b4 Then
                    intDir = 4
                'Right
                ElseIf b3 Then
                    intDir = 3
                'Left
                ElseIf b2 Then
                    intDir = 2
                'Up
                ElseIf b1 Then
                    intDir = 1
                End If
            End If
        End If
    
    'Set Direction Changes
        'Up
        If intDir = 1 Then
            AI_Object.nDir = 1
            AI_Object.Vx = 0
            AI_Object.Vy = -1
            pDirLast = 1
            pDirPrevent = 4
        'Left
        ElseIf intDir = 2 Then
            AI_Object.nDir = 2
            AI_Object.Vx = -1
            AI_Object.Vy = 0
            pDirLast = 2
            pDirPrevent = 3
        'Right
        ElseIf intDir = 3 Then
            AI_Object.nDir = 3
            AI_Object.Vx = 1
            AI_Object.Vy = 0
            pDirLast = 3
            pDirPrevent = 2
        'Down
        ElseIf intDir = 4 Then
            AI_Object.nDir = 4
            AI_Object.Vx = 0
            AI_Object.Vy = 1
            pDirLast = 4
            pDirPrevent = 1
        End If

End Sub

Public Sub ChangeSpeedCheck(ByRef AI_Object As Sprite_Ghosts)
'Changes speed of Ghost safely to prevent moving out of track grid.

    Dim bXok As Boolean 'True: x-coord is in a position allowing for speed change
    Dim bYok As Boolean 'True: y-coord is in a position allowing for speed change
    
    'initialize booleans
    bXok = False
    bYok = False
    
    'Chase
    If pMoveType = 1 Then
        'proper x coordinate
        If ((AI_Object.Xcol - 21) Mod 4) = 0 Then: bXok = True
        'proper y coordinate
        If ((AI_Object.Yrow - 29) Mod 4) = 0 Then: bYok = True
        'change speed
        If bXok And bYok Then: AI_Object.Speed = 4
    'Scatter
    ElseIf pMoveType = 2 Then
        'proper x coordinate
        If ((AI_Object.Xcol - 21) Mod 4) = 0 Then: bXok = True
        'proper y coordinate
        If ((AI_Object.Yrow - 29) Mod 4) = 0 Then: bYok = True
        'change speed
        If bXok And bYok Then: AI_Object.Speed = 4
    'Scared
    ElseIf pMoveType = 3 Then
        'proper x coordinate
        If ((AI_Object.Xcol - 21) Mod 2) = 0 Then: bXok = True
        'proper y coordinate
        If ((AI_Object.Yrow - 29) Mod 2) = 0 Then: bYok = True
        'change speed
        If bXok And bYok Then: AI_Object.Speed = 2
    'Eyes
    ElseIf pMoveType = 4 Then
        'proper x coordinate
        If ((AI_Object.Xcol - 21) Mod 8) = 0 Then: bXok = True
        'proper y coordinate
        If ((AI_Object.Yrow - 29) Mod 8) = 0 Then: bYok = True
        'change speed
        If bXok And bYok Then: AI_Object.Speed = 8
    End If

End Sub

Public Sub GhostCageAnimations(ByRef AI_Object As Sprite_Ghosts)
'Animations for Ghosts in the Ghost Cage (no movement track within cage)

    'Prevent This Anim if Ghost is in Hidden Location
    If Not (AI_Object.Xcol = 1 And AI_Object.Yrow = 253) Then
        
        'Enter
        If pCageMoveType = 1 Then
            
            'Blinky or Pinky
            If AI_Object.GhostName = "Blinky" Or AI_Object.GhostName = "Pinky" Then
                If pCageMoveCounter = 1 Then
                    AI_Object.Xcol = 121 'cave entrance
                    AI_Object.Yrow = 109 'cave entrance
                    pCageMoveCounter = 2
                ElseIf pCageMoveCounter = 2 Then
                    AI_Object.Xcol = AI_Object.Xcol
                    AI_Object.Yrow = AI_Object.Yrow + 8
                    AI_Object.nDir = 4 'down
                    pCageMoveCounter = 3
                ElseIf pCageMoveCounter = 3 Then
                    AI_Object.Xcol = AI_Object.Xcol
                    AI_Object.Yrow = AI_Object.Yrow + 8
                    pCageMoveCounter = 4
                ElseIf pCageMoveCounter = 4 Then
                    AI_Object.Xcol = AI_Object.Xcol
                    AI_Object.Yrow = AI_Object.Yrow + 8
                    pCageMoveCounter = 1 'reset
                    pCageMoveType = 3 'exit
                    AI_Object.Eaten = False 'allow MsP collisions
                    AI_Object.Points = 200 'reset points
                    'Ghost Mode and Scare State Timers
                    If Not AI_Object.ScaredStateAgain Then
                        AI_Object.GhostMode = 1 'regular ghost anim
                    ElseIf AI_Object.ScaredStateAgain Then
                        AI_Object.ScaredStateAgain = False
                        AI_Object.GhostMode = 2 'blue scared ghost anim
                        AI_Object.MatchBoolean = True 'link scared state up with the rest of the ghosts
                    End If
                    'Audio
                    If (Blinky.GhostMode = 2) Or (Blinky.GhostMode = 3) Or (Pinky.GhostMode = 2) Or (Pinky.GhostMode = 3) Or (Inky.GhostMode = 2) Or (Inky.GhostMode = 3) Or (Sue.GhostMode = 2) Or (Sue.GhostMode = 3) Then
                        Call PlayAudio(arrPowerDotAmbient, False) 'play power dot ambient audio
                    End If
                    If (Blinky.GhostMode <> 4) Or (Pinky.GhostMode <> 4) Or (Inky.GhostMode <> 4) Or (Sue.GhostMode <> 4) Then
                        Call StopAudio(arrGhostEyes, False) 'stop ghost eyes audio
                    End If
                End If
            
            'Inky
            ElseIf AI_Object.GhostName = "Inky" Then
                If pCageMoveCounter = 1 Then
                    AI_Object.Xcol = 121 'cave entrance
                    AI_Object.Yrow = 109 'cave entrance
                    pCageMoveCounter = 2
                ElseIf pCageMoveCounter = 2 Then
                    AI_Object.Xcol = AI_Object.Xcol
                    AI_Object.Yrow = AI_Object.Yrow + 8
                    AI_Object.nDir = 4 'down
                    pCageMoveCounter = 3
                ElseIf pCageMoveCounter = 3 Then
                    AI_Object.Xcol = AI_Object.Xcol
                    AI_Object.Yrow = AI_Object.Yrow + 8
                    pCageMoveCounter = 4
                ElseIf pCageMoveCounter = 4 Then
                    AI_Object.Xcol = AI_Object.Xcol
                    AI_Object.Yrow = AI_Object.Yrow + 8
                    pCageMoveCounter = 5
                ElseIf pCageMoveCounter = 5 Then
                    AI_Object.Xcol = AI_Object.Xcol - 8
                    AI_Object.Yrow = AI_Object.Yrow
                    AI_Object.nDir = 2 'left
                    pCageMoveCounter = 6
                ElseIf pCageMoveCounter = 6 Then
                    AI_Object.Xcol = AI_Object.Xcol - 8
                    AI_Object.Yrow = AI_Object.Yrow
                    pCageMoveCounter = 1 'reset
                    pCageMoveType = 3 'exit
                    AI_Object.Eaten = False 'allow MsP collisions
                    AI_Object.Points = 200 'reset points
                    'Ghost Mode and Scare State Timers
                    If Not AI_Object.ScaredStateAgain Then
                        AI_Object.GhostMode = 1 'regular ghost anim
                    ElseIf AI_Object.ScaredStateAgain Then
                        AI_Object.ScaredStateAgain = False
                        AI_Object.GhostMode = 2 'blue scared ghost anim
                        AI_Object.MatchBoolean = True 'link scared state up with the rest of the ghosts
                    End If
                    'Audio
                    If (Blinky.GhostMode = 2) Or (Blinky.GhostMode = 3) Or (Pinky.GhostMode = 2) Or (Pinky.GhostMode = 3) Or (Inky.GhostMode = 2) Or (Inky.GhostMode = 3) Or (Sue.GhostMode = 2) Or (Sue.GhostMode = 3) Then
                        Call PlayAudio(arrPowerDotAmbient, False) 'play power dot ambient audio
                    End If
                    If (Blinky.GhostMode <> 4) Or (Pinky.GhostMode <> 4) Or (Inky.GhostMode <> 4) Or (Sue.GhostMode <> 4) Then
                        Call StopAudio(arrGhostEyes, False) 'stop ghost eyes audio
                    End If
                End If
            
            'Sue
            ElseIf AI_Object.GhostName = "Sue" Then
                If pCageMoveCounter = 1 Then
                    AI_Object.Xcol = 121 'cave entrance
                    AI_Object.Yrow = 109 'cave entrance
                    pCageMoveCounter = 2
                ElseIf pCageMoveCounter = 2 Then
                    AI_Object.Xcol = AI_Object.Xcol
                    AI_Object.Yrow = AI_Object.Yrow + 8
                    AI_Object.nDir = 4 'down
                    pCageMoveCounter = 3
                ElseIf pCageMoveCounter = 3 Then
                    AI_Object.Xcol = AI_Object.Xcol
                    AI_Object.Yrow = AI_Object.Yrow + 8
                    pCageMoveCounter = 4
                ElseIf pCageMoveCounter = 4 Then
                    AI_Object.Xcol = AI_Object.Xcol
                    AI_Object.Yrow = AI_Object.Yrow + 8
                    pCageMoveCounter = 5
                ElseIf pCageMoveCounter = 5 Then
                    AI_Object.Xcol = AI_Object.Xcol + 8
                    AI_Object.Yrow = AI_Object.Yrow
                    AI_Object.nDir = 3 'right
                    pCageMoveCounter = 6
                ElseIf pCageMoveCounter = 6 Then
                    AI_Object.Xcol = AI_Object.Xcol + 8
                    AI_Object.Yrow = AI_Object.Yrow
                    pCageMoveCounter = 1 'reset
                    pCageMoveType = 3 'exit
                    AI_Object.Eaten = False 'allow MsP collisions
                    AI_Object.Points = 200 'reset points
                    'Ghost Mode and Scare State Timers
                    If Not AI_Object.ScaredStateAgain Then
                        AI_Object.GhostMode = 1 'regular ghost anim
                    ElseIf AI_Object.ScaredStateAgain Then
                        AI_Object.ScaredStateAgain = False
                        AI_Object.GhostMode = 2 'blue scared ghost anim
                        AI_Object.MatchBoolean = True 'link scared state up with the rest of the ghosts
                    End If
                    'Audio
                    If (Blinky.GhostMode = 2) Or (Blinky.GhostMode = 3) Or (Pinky.GhostMode = 2) Or (Pinky.GhostMode = 3) Or (Inky.GhostMode = 2) Or (Inky.GhostMode = 3) Or (Sue.GhostMode = 2) Or (Sue.GhostMode = 3) Then
                        Call PlayAudio(arrPowerDotAmbient, False) 'play power dot ambient audio
                    End If
                    If (Blinky.GhostMode <> 4) Or (Pinky.GhostMode <> 4) Or (Inky.GhostMode <> 4) Or (Sue.GhostMode <> 4) Then
                        Call StopAudio(arrGhostEyes, False) 'stop ghost eyes audio
                    End If
                End If
            End If
        
        
        'Wiggle
        ElseIf pCageMoveType = 2 Then
            
            'Blinky is never in this movement type
            
            'Pinky
            If AI_Object.GhostName = "Pinky" Then
                If pCageMoveCounter = 0 Then 'prevents first frame movement
                    pCageMoveCounter = 1
                ElseIf pCageMoveCounter = 1 Then
                    AI_Object.Xcol = AI_Object.Xcol
                    AI_Object.Yrow = AI_Object.Yrow + 2
                    AI_Object.nDir = 4 'down
                    pCageMoveCounter = 2
                ElseIf pCageMoveCounter = 2 Then
                    AI_Object.Xcol = AI_Object.Xcol
                    AI_Object.Yrow = AI_Object.Yrow + 2
                    pCageMoveCounter = 3
                ElseIf pCageMoveCounter = 3 Then
                    AI_Object.Xcol = AI_Object.Xcol
                    AI_Object.Yrow = AI_Object.Yrow - 2
                    AI_Object.nDir = 1 'up
                    pCageMoveCounter = 4
                ElseIf pCageMoveCounter = 4 Then
                    AI_Object.Xcol = AI_Object.Xcol
                    AI_Object.Yrow = AI_Object.Yrow - 2
                    pCageMoveCounter = 5
                ElseIf pCageMoveCounter = 5 Then
                    AI_Object.Xcol = AI_Object.Xcol
                    AI_Object.Yrow = AI_Object.Yrow - 2
                    pCageMoveCounter = 6
                ElseIf pCageMoveCounter = 6 Then
                    AI_Object.Xcol = AI_Object.Xcol
                    AI_Object.Yrow = AI_Object.Yrow - 2
                    pCageMoveCounter = 7
                ElseIf pCageMoveCounter = 7 Then
                    AI_Object.Xcol = AI_Object.Xcol
                    AI_Object.Yrow = AI_Object.Yrow + 2
                    AI_Object.nDir = 4 'down
                    pCageMoveCounter = 8
                ElseIf pCageMoveCounter = 8 Then
                    AI_Object.Xcol = AI_Object.Xcol
                    AI_Object.Yrow = AI_Object.Yrow + 2
                    pCageMoveCounter = 1
                End If
            
            'Inky or Sue
            ElseIf AI_Object.GhostName = "Inky" Or AI_Object.GhostName = "Sue" Then
                If pCageMoveCounter = 0 Then 'prevents first frame movement
                    pCageMoveCounter = 1
                ElseIf pCageMoveCounter = 1 Then
                    AI_Object.Xcol = AI_Object.Xcol
                    AI_Object.Yrow = AI_Object.Yrow - 2
                    AI_Object.nDir = 1 'up
                    pCageMoveCounter = 2
                ElseIf pCageMoveCounter = 2 Then
                    AI_Object.Xcol = AI_Object.Xcol
                    AI_Object.Yrow = AI_Object.Yrow - 2
                    pCageMoveCounter = 3
                ElseIf pCageMoveCounter = 3 Then
                    AI_Object.Xcol = AI_Object.Xcol
                    AI_Object.Yrow = AI_Object.Yrow + 2
                    AI_Object.nDir = 4 'down
                    pCageMoveCounter = 4
                ElseIf pCageMoveCounter = 4 Then
                    AI_Object.Xcol = AI_Object.Xcol
                    AI_Object.Yrow = AI_Object.Yrow + 2
                    pCageMoveCounter = 5
                ElseIf pCageMoveCounter = 5 Then
                    AI_Object.Xcol = AI_Object.Xcol
                    AI_Object.Yrow = AI_Object.Yrow + 2
                    pCageMoveCounter = 6
                ElseIf pCageMoveCounter = 6 Then
                    AI_Object.Xcol = AI_Object.Xcol
                    AI_Object.Yrow = AI_Object.Yrow + 2
                    pCageMoveCounter = 7
                ElseIf pCageMoveCounter = 7 Then
                    AI_Object.Xcol = AI_Object.Xcol
                    AI_Object.Yrow = AI_Object.Yrow - 2
                    AI_Object.nDir = 1 'up
                    pCageMoveCounter = 8
                ElseIf pCageMoveCounter = 8 Then
                    AI_Object.Xcol = AI_Object.Xcol
                    AI_Object.Yrow = AI_Object.Yrow - 2
                    pCageMoveCounter = 1
                End If
            End If
        
        
        'Exit
        ElseIf pCageMoveType = 3 Then
            
            'Blinky or Pinky
            If AI_Object.GhostName = "Blinky" Or AI_Object.GhostName = "Pinky" Then
                If pCageMoveCounter = 1 Then
                    AI_Object.Xcol = AI_Object.XcolGC 'Start Exit from GC coords
                    AI_Object.Yrow = AI_Object.YrowGC 'Start Exit from GC coords
                    pCageMoveCounter = 2
                ElseIf pCageMoveCounter = 2 Then
                    AI_Object.Xcol = AI_Object.Xcol
                    AI_Object.Yrow = AI_Object.Yrow - 2
                    AI_Object.nDir = 1 'up
                    pCageMoveCounter = 3
                ElseIf pCageMoveCounter = 3 Then
                    AI_Object.Xcol = AI_Object.Xcol
                    AI_Object.Yrow = AI_Object.Yrow - 2
                    pCageMoveCounter = 4
                ElseIf pCageMoveCounter = 4 Then
                    AI_Object.Xcol = AI_Object.Xcol
                    AI_Object.Yrow = AI_Object.Yrow - 2
                    pCageMoveCounter = 5
                ElseIf pCageMoveCounter = 5 Then
                    AI_Object.Xcol = AI_Object.Xcol
                    AI_Object.Yrow = AI_Object.Yrow - 2
                    pCageMoveCounter = 6
                ElseIf pCageMoveCounter = 6 Then
                    AI_Object.Xcol = AI_Object.Xcol
                    AI_Object.Yrow = AI_Object.Yrow - 2
                    pCageMoveCounter = 7
                ElseIf pCageMoveCounter = 7 Then
                    AI_Object.Xcol = AI_Object.Xcol
                    AI_Object.Yrow = AI_Object.Yrow - 2
                    pCageMoveCounter = 8
                ElseIf pCageMoveCounter = 8 Then
                    AI_Object.Xcol = AI_Object.Xcol
                    AI_Object.Yrow = AI_Object.Yrow - 2
                    pCageMoveCounter = 9
                ElseIf pCageMoveCounter = 9 Then
                    AI_Object.Xcol = AI_Object.Xcol
                    AI_Object.Yrow = AI_Object.Yrow - 2
                    pCageMoveCounter = 10
                ElseIf pCageMoveCounter = 10 Then
                    AI_Object.Xcol = AI_Object.Xcol
                    AI_Object.Yrow = AI_Object.Yrow - 2
                    pCageMoveCounter = 11
                ElseIf pCageMoveCounter = 11 Then
                    AI_Object.Xcol = AI_Object.Xcol
                    AI_Object.Yrow = AI_Object.Yrow - 2
                    pCageMoveCounter = 12
                ElseIf pCageMoveCounter = 12 Then
                    AI_Object.Xcol = AI_Object.Xcol
                    AI_Object.Yrow = AI_Object.Yrow - 2
                    pCageMoveCounter = 13
                ElseIf pCageMoveCounter = 13 Then
                    AI_Object.Xcol = AI_Object.Xcol
                    AI_Object.Yrow = AI_Object.Yrow - 2
                    pCageMoveCounter = 0
                    pCageMoveType = 0 'not in cage
                    'change move type to phase type
                    If AI_Object.GhostMode = 1 Then
                        pMoveType = GhostPhaseType 'Chase or Scatter
                    Else
                        pMoveType = 3 'Scared
                    End If
                End If
            
            'Inky
            ElseIf AI_Object.GhostName = "Inky" Then
                If pCageMoveCounter = 1 Then
                    AI_Object.Xcol = AI_Object.XcolGC 'Start Exit from GC coords
                    AI_Object.Yrow = AI_Object.YrowGC 'Start Exit from GC coords
                    pCageMoveCounter = 2
                ElseIf pCageMoveCounter = 2 Then
                    AI_Object.Xcol = AI_Object.Xcol + 2
                    AI_Object.Yrow = AI_Object.Yrow
                    AI_Object.nDir = 3 'right
                    pCageMoveCounter = 3
                ElseIf pCageMoveCounter = 3 Then
                    AI_Object.Xcol = AI_Object.Xcol + 2
                    AI_Object.Yrow = AI_Object.Yrow
                    pCageMoveCounter = 4
                ElseIf pCageMoveCounter = 4 Then
                    AI_Object.Xcol = AI_Object.Xcol + 2
                    AI_Object.Yrow = AI_Object.Yrow
                    pCageMoveCounter = 5
                ElseIf pCageMoveCounter = 5 Then
                    AI_Object.Xcol = AI_Object.Xcol + 2
                    AI_Object.Yrow = AI_Object.Yrow
                    pCageMoveCounter = 6
                ElseIf pCageMoveCounter = 6 Then
                    AI_Object.Xcol = AI_Object.Xcol + 2
                    AI_Object.Yrow = AI_Object.Yrow
                    pCageMoveCounter = 7
                ElseIf pCageMoveCounter = 7 Then
                    AI_Object.Xcol = AI_Object.Xcol + 2
                    AI_Object.Yrow = AI_Object.Yrow
                    pCageMoveCounter = 8
                ElseIf pCageMoveCounter = 8 Then
                    AI_Object.Xcol = AI_Object.Xcol + 2
                    AI_Object.Yrow = AI_Object.Yrow
                    pCageMoveCounter = 9
                ElseIf pCageMoveCounter = 9 Then
                    AI_Object.Xcol = AI_Object.Xcol + 2
                    AI_Object.Yrow = AI_Object.Yrow
                    pCageMoveCounter = 10
                ElseIf pCageMoveCounter = 10 Then
                    AI_Object.Xcol = AI_Object.Xcol
                    AI_Object.Yrow = AI_Object.Yrow - 2
                    AI_Object.nDir = 1 'up
                    pCageMoveCounter = 11
                ElseIf pCageMoveCounter = 11 Then
                    AI_Object.Xcol = AI_Object.Xcol
                    AI_Object.Yrow = AI_Object.Yrow - 2
                    pCageMoveCounter = 12
                ElseIf pCageMoveCounter = 12 Then
                    AI_Object.Xcol = AI_Object.Xcol
                    AI_Object.Yrow = AI_Object.Yrow - 2
                    pCageMoveCounter = 13
                ElseIf pCageMoveCounter = 13 Then
                    AI_Object.Xcol = AI_Object.Xcol
                    AI_Object.Yrow = AI_Object.Yrow - 2
                    pCageMoveCounter = 14
                ElseIf pCageMoveCounter = 14 Then
                    AI_Object.Xcol = AI_Object.Xcol
                    AI_Object.Yrow = AI_Object.Yrow - 2
                    pCageMoveCounter = 15
                ElseIf pCageMoveCounter = 15 Then
                    AI_Object.Xcol = AI_Object.Xcol
                    AI_Object.Yrow = AI_Object.Yrow - 2
                    pCageMoveCounter = 16
                ElseIf pCageMoveCounter = 16 Then
                    AI_Object.Xcol = AI_Object.Xcol
                    AI_Object.Yrow = AI_Object.Yrow - 2
                    pCageMoveCounter = 17
                ElseIf pCageMoveCounter = 17 Then
                    AI_Object.Xcol = AI_Object.Xcol
                    AI_Object.Yrow = AI_Object.Yrow - 2
                    pCageMoveCounter = 18
                ElseIf pCageMoveCounter = 18 Then
                    AI_Object.Xcol = AI_Object.Xcol
                    AI_Object.Yrow = AI_Object.Yrow - 2
                    pCageMoveCounter = 19
                ElseIf pCageMoveCounter = 19 Then
                    AI_Object.Xcol = AI_Object.Xcol
                    AI_Object.Yrow = AI_Object.Yrow - 2
                    pCageMoveCounter = 20
                ElseIf pCageMoveCounter = 20 Then
                    AI_Object.Xcol = AI_Object.Xcol
                    AI_Object.Yrow = AI_Object.Yrow - 2
                    pCageMoveCounter = 21
                ElseIf pCageMoveCounter = 21 Then
                    AI_Object.Xcol = AI_Object.Xcol
                    AI_Object.Yrow = AI_Object.Yrow - 2
                    pCageMoveCounter = 0
                    pCageMoveType = 0 'not in cage
                    'change move type to phase type
                    If AI_Object.GhostMode = 1 Then
                        pMoveType = GhostPhaseType 'Chase or Scatter
                    Else
                        pMoveType = 3 'Scared
                    End If
                End If
            
            'Sue
            ElseIf AI_Object.GhostName = "Sue" Then
                If pCageMoveCounter = 1 Then
                    AI_Object.Xcol = AI_Object.XcolGC 'Start Exit from GC coords
                    AI_Object.Yrow = AI_Object.YrowGC 'Start Exit from GC coords
                    pCageMoveCounter = 2
                ElseIf pCageMoveCounter = 2 Then
                    AI_Object.Xcol = AI_Object.Xcol - 2
                    AI_Object.Yrow = AI_Object.Yrow
                    AI_Object.nDir = 2 'left
                    pCageMoveCounter = 3
                ElseIf pCageMoveCounter = 3 Then
                    AI_Object.Xcol = AI_Object.Xcol - 2
                    AI_Object.Yrow = AI_Object.Yrow
                    pCageMoveCounter = 4
                ElseIf pCageMoveCounter = 4 Then
                    AI_Object.Xcol = AI_Object.Xcol - 2
                    AI_Object.Yrow = AI_Object.Yrow
                    pCageMoveCounter = 5
                ElseIf pCageMoveCounter = 5 Then
                    AI_Object.Xcol = AI_Object.Xcol - 2
                    AI_Object.Yrow = AI_Object.Yrow
                    pCageMoveCounter = 6
                ElseIf pCageMoveCounter = 6 Then
                    AI_Object.Xcol = AI_Object.Xcol - 2
                    AI_Object.Yrow = AI_Object.Yrow
                    pCageMoveCounter = 7
                ElseIf pCageMoveCounter = 7 Then
                    AI_Object.Xcol = AI_Object.Xcol - 2
                    AI_Object.Yrow = AI_Object.Yrow
                    pCageMoveCounter = 8
                ElseIf pCageMoveCounter = 8 Then
                    AI_Object.Xcol = AI_Object.Xcol - 2
                    AI_Object.Yrow = AI_Object.Yrow
                    pCageMoveCounter = 9
                ElseIf pCageMoveCounter = 9 Then
                    AI_Object.Xcol = AI_Object.Xcol - 2
                    AI_Object.Yrow = AI_Object.Yrow
                    pCageMoveCounter = 10
                ElseIf pCageMoveCounter = 10 Then
                    AI_Object.Xcol = AI_Object.Xcol
                    AI_Object.Yrow = AI_Object.Yrow - 2
                    AI_Object.nDir = 1 'up
                    pCageMoveCounter = 11
                ElseIf pCageMoveCounter = 11 Then
                    AI_Object.Xcol = AI_Object.Xcol
                    AI_Object.Yrow = AI_Object.Yrow - 2
                    pCageMoveCounter = 12
                ElseIf pCageMoveCounter = 12 Then
                    AI_Object.Xcol = AI_Object.Xcol
                    AI_Object.Yrow = AI_Object.Yrow - 2
                    pCageMoveCounter = 13
                ElseIf pCageMoveCounter = 13 Then
                    AI_Object.Xcol = AI_Object.Xcol
                    AI_Object.Yrow = AI_Object.Yrow - 2
                    pCageMoveCounter = 14
                ElseIf pCageMoveCounter = 14 Then
                    AI_Object.Xcol = AI_Object.Xcol
                    AI_Object.Yrow = AI_Object.Yrow - 2
                    pCageMoveCounter = 15
                ElseIf pCageMoveCounter = 15 Then
                    AI_Object.Xcol = AI_Object.Xcol
                    AI_Object.Yrow = AI_Object.Yrow - 2
                    pCageMoveCounter = 16
                ElseIf pCageMoveCounter = 16 Then
                    AI_Object.Xcol = AI_Object.Xcol
                    AI_Object.Yrow = AI_Object.Yrow - 2
                    pCageMoveCounter = 17
                ElseIf pCageMoveCounter = 17 Then
                    AI_Object.Xcol = AI_Object.Xcol
                    AI_Object.Yrow = AI_Object.Yrow - 2
                    pCageMoveCounter = 18
                ElseIf pCageMoveCounter = 18 Then
                    AI_Object.Xcol = AI_Object.Xcol
                    AI_Object.Yrow = AI_Object.Yrow - 2
                    pCageMoveCounter = 19
                ElseIf pCageMoveCounter = 19 Then
                    AI_Object.Xcol = AI_Object.Xcol
                    AI_Object.Yrow = AI_Object.Yrow - 2
                    pCageMoveCounter = 20
                ElseIf pCageMoveCounter = 20 Then
                    AI_Object.Xcol = AI_Object.Xcol
                    AI_Object.Yrow = AI_Object.Yrow - 2
                    pCageMoveCounter = 21
                ElseIf pCageMoveCounter = 21 Then
                    AI_Object.Xcol = AI_Object.Xcol
                    AI_Object.Yrow = AI_Object.Yrow - 2
                    pCageMoveCounter = 0
                    pCageMoveType = 0 'not in cage
                    'change move type to phase type
                    If AI_Object.GhostMode = 1 Then
                        pMoveType = GhostPhaseType 'Chase or Scatter
                    Else
                        pMoveType = 3 'Scared
                    End If
                End If
            End If
        End If
    End If


End Sub

Public Sub TimerCageReset(ByVal strGhostName As String)
'Reset Cage Timer

    'Cage Timer
        'Blinky
        If strGhostName = "Blinky" Then
            pTimerCage = 10 'frames
        'Pinky
        ElseIf strGhostName = "Pinky" Then
            pTimerCage = 60 'frames
        'Inky
        ElseIf strGhostName = "Inky" Then
            pTimerCage = 90 'frames
        'Sue
        ElseIf strGhostName = "Sue" Then
            pTimerCage = 120 'frames
        End If

End Sub

Public Sub TimerCageDecrement()
'Count Cage Timer down

    'Cage Timer
    If pMoveType = 0 Then
        pTimerCage = pTimerCage - 1
    End If

End Sub
