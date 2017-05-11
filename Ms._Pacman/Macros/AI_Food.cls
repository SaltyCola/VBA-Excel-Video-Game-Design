VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "AI_Food"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'=================================================================================='
'=================================== Properties ==================================='
'=================================================================================='

'Timer
Private pTimer As Integer 'Tells Food when to enter the map (10 frames ~ 1 seconds)

'Circular Motion
Private pCircularMotion As Integer 'Determines which direction to go around the center loop. [1=Clockwise ; 2=CounterClockwise]

'Entrance and Exit Tunnels [1-4 or 1-2 (numbered left to right, top to bottom)]
Private pTunnelEnter As Integer 'Corner from where food will enter the map
Private pTunnelExit As Integer 'Corner from where food will exit the map

'Object Coordinates
Private pXobject As Integer 'object's x or column coordinate
Private pYobject As Integer 'object's y or row coordinate

'Target Coordinates
Private pXtarget As Integer 'target's x or column coordinate
Private pYtarget As Integer 'target's y or row coordinate
Private pTargetCounter As Integer 'Counts how many targets have been chosen so that final target is exit tunnel

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
' Entrance Timer
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Property Get Timer() As Integer
    Timer = pTimer
End Property
Public Property Let Timer(Value As Integer)
    pTimer = Value
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
' Entrance Tunnel
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Property Get TunnelEnter() As Integer
    TunnelEnter = pTunnelEnter
End Property
Public Property Let TunnelEnter(Value As Integer)
    pTunnelEnter = Value
End Property

'''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
' Exit Tunnel
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Property Get TunnelExit() As Integer
    TunnelExit = pTunnelExit
End Property
Public Property Let TunnelExit(Value As Integer)
    pTunnelExit = Value
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
' Target Counter
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Property Get TargetCounter() As Integer
    TargetCounter = pTargetCounter
End Property
Public Property Let TargetCounter(Value As Integer)
    pTargetCounter = Value
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

Public Sub TargetChange()
'Change coordinates of the target

    'Increment Target Counter
        pTargetCounter = pTargetCounter + 1
    
    'Find Next Target
        'Exit Tunnel
        If pTargetCounter = 7 Then
            'tunnel 1
            If pTunnelExit = 1 Then
                pXtarget = TunnelCoords(1, 1)
                pYtarget = TunnelCoords(1, 2)
            'tunnel 2
            ElseIf pTunnelExit = 2 Then
                pXtarget = TunnelCoords(2, 1)
                pYtarget = TunnelCoords(2, 2)
            'tunnel 3
            ElseIf pTunnelExit = 3 Then
                pXtarget = TunnelCoords(3, 1)
                pYtarget = TunnelCoords(3, 2)
            'tunnel 4
            ElseIf pTunnelExit = 4 Then
                pXtarget = TunnelCoords(4, 1)
                pYtarget = TunnelCoords(4, 2)
            End If
        'Pre Exit Tunnel Target
        ElseIf pTargetCounter = 6 Then
            'tunnel 1
            If pTunnelExit = 1 Then
                pXtarget = TunnelCoords(1, 1) + 36 '36 puts target at entrance to exit tunnel
                pYtarget = TunnelCoords(1, 2)
            'tunnel 2
            ElseIf pTunnelExit = 2 Then
                pXtarget = TunnelCoords(2, 1) - 36 '36 puts target at entrance to exit tunnel
                pYtarget = TunnelCoords(2, 2)
            'tunnel 3
            ElseIf pTunnelExit = 3 Then
                pXtarget = TunnelCoords(3, 1) + 36 '36 puts target at entrance to exit tunnel
                pYtarget = TunnelCoords(3, 2)
            'tunnel 4
            ElseIf pTunnelExit = 4 Then
                pXtarget = TunnelCoords(4, 1) - 36 '36 puts target at entrance to exit tunnel
                pYtarget = TunnelCoords(4, 2)
            End If
        'corner 1 to (2 or 3)
        ElseIf pXtarget = CenterCornerCoords(1, 1) And pYtarget = CenterCornerCoords(1, 2) Then
            'Clockwise
            If pCircularMotion = 1 Then
                pXtarget = CenterCornerCoords(2, 1)
                pYtarget = CenterCornerCoords(2, 2)
            'CounterClockwise
            ElseIf pCircularMotion = 2 Then
                pXtarget = CenterCornerCoords(3, 1)
                pYtarget = CenterCornerCoords(3, 2)
            End If
        'corner 2 to (4 or 1)
        ElseIf pXtarget = CenterCornerCoords(2, 1) And pYtarget = CenterCornerCoords(2, 2) Then
            'Clockwise
            If pCircularMotion = 1 Then
                pXtarget = CenterCornerCoords(4, 1)
                pYtarget = CenterCornerCoords(4, 2)
            'CounterClockwise
            ElseIf pCircularMotion = 2 Then
                pXtarget = CenterCornerCoords(1, 1)
                pYtarget = CenterCornerCoords(1, 2)
            End If
        'corner 3 to (1 or 4)
        ElseIf pXtarget = CenterCornerCoords(3, 1) And pYtarget = CenterCornerCoords(3, 2) Then
            'Clockwise
            If pCircularMotion = 1 Then
                pXtarget = CenterCornerCoords(1, 1)
                pYtarget = CenterCornerCoords(1, 2)
            'CounterClockwise
            ElseIf pCircularMotion = 2 Then
                pXtarget = CenterCornerCoords(4, 1)
                pYtarget = CenterCornerCoords(4, 2)
            End If
        'corner 4 to (3 or 2)
        ElseIf pXtarget = CenterCornerCoords(4, 1) And pYtarget = CenterCornerCoords(4, 2) Then
            'Clockwise
            If pCircularMotion = 1 Then
                pXtarget = CenterCornerCoords(3, 1)
                pYtarget = CenterCornerCoords(3, 2)
            'CounterClockwise
            ElseIf pCircularMotion = 2 Then
                pXtarget = CenterCornerCoords(2, 1)
                pYtarget = CenterCornerCoords(2, 2)
            End If
        End If

End Sub

Public Sub CalculateSightLines(ByRef AI_Object As Sprite_Food)
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

Public Sub DetermineDirection(ByRef AI_Object As Sprite_Food, ByVal bUp As Boolean, ByVal bLeft As Boolean, ByVal bRight As Boolean, ByVal bDown As Boolean)
'Determines which direction the food object should take when it reaches an intersection.

    Dim intRnd As Integer 'random choice for two correct decisions (will never be more than 2)
    Dim intDir As Integer 'Direction of next turn [1=up ; 2=left ; 3=right ; 4=down]
    Dim b1 As Boolean 'up is a correct choice
    Dim b2 As Boolean 'left is a correct choice
    Dim b3 As Boolean 'right is a correct choice
    Dim b4 As Boolean 'down is a correct choice
    Dim sl_short As Double 'shortest distance for all 4 sightlines
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
        'Initialize to value higher than possible
        sl_short = 5000
        'Up
        If pSLup < sl_short And bUp And pDirPrevent <> 1 Then: sl_short = pSLup
        'Left
        If pSLleft < sl_short And bLeft And pDirPrevent <> 2 Then: sl_short = pSLleft
        'Right
        If pSLright < sl_short And bRight And pDirPrevent <> 3 Then: sl_short = pSLright
        'Down
        If pSLdown < sl_short And bDown And pDirPrevent <> 4 Then: sl_short = pSLdown
    
    'Get Correct Choices
        'Up
        If pSLup = sl_short And bUp And pDirPrevent <> 1 Then
            b1 = True
            intDir = 1
            cc = cc + 1
        End If
        'Left
        If pSLleft = sl_short And bLeft And pDirPrevent <> 2 Then
            b2 = True
            intDir = 2
            cc = cc + 1
        End If
        'Right
        If pSLright = sl_short And bRight And pDirPrevent <> 3 Then
            b3 = True
            intDir = 3
            cc = cc + 1
        End If
        'Down
        If pSLdown = sl_short And bDown And pDirPrevent <> 4 Then
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

Public Sub TimerReset()
'Reset Timer

    'Timer [200 frames ~= 20 seconds]
    pTimer = 200 '200 'frames
    
    'Reset Target Counter
    pTargetCounter = 0

End Sub

Public Sub TimerDecrement()
'Count Timer down

    pTimer = pTimer - 1

End Sub

Public Sub TunnelsRandom(ByRef AI_Object As Sprite_Food)
'Randomly select an entrance and exit tunnel. Entrance tunnel also determines
    'first target location (closest central corner), and first "pDirLast" & "pDirPrevent"

    Dim rndEnter As Integer 'random choice for entrance tunnel
    Dim rndExit As Integer 'random choice for exit tunnel
    Dim rndChange As Integer 'random choice for changing exit tunnel in case of pTunnelEnter = pTunnelExit
    
    'Random Choice
        'Tunnel Enter
            Randomize
            rndEnter = Int((100 - 1 + 1) * Rnd + 1) 'int((upperbound-lowerbound+1)*Rnd+lowerbound)
        'Tunnel Exit
            Randomize
            rndExit = Int((100 - 1 + 1) * Rnd + 1) 'int((upperbound-lowerbound+1)*Rnd+lowerbound)
    
    '4 Tunnel Options
    If NumberOfTunnels = 4 Then
        'Tunnel Enter
            'Choice 1
            If rndEnter <= 25 Then
                pTunnelEnter = 1
            'Choice 2
            ElseIf rndEnter > 25 And rndEnter <= 50 Then
                pTunnelEnter = 2
            'Choice 3
            ElseIf rndEnter > 50 And rndEnter <= 75 Then
                pTunnelEnter = 3
            'Choice 4
            ElseIf rndEnter > 75 Then
                pTunnelEnter = 4
            End If
        'Tunnel Exit
            'Choice 1
            If rndExit <= 25 Then
                pTunnelExit = 1
            'Choice 2
            ElseIf rndExit > 25 And rndExit <= 50 Then
                pTunnelExit = 2
            'Choice 3
            ElseIf rndExit > 50 And rndExit <= 75 Then
                pTunnelExit = 3
            'Choice 4
            ElseIf rndExit > 75 Then
                pTunnelExit = 4
            End If
    
    '2 Tunnel Options
    ElseIf NumberOfTunnels = 2 Then
        'Tunnel Enter
            'Choice 1
            If rndEnter <= 50 Then
                pTunnelEnter = 1
            'Choice 2
            ElseIf rndEnter > 50 Then
                pTunnelEnter = 2
            End If
        'Tunnel Exit
            'Choice 1
            If rndExit <= 50 Then
                pTunnelExit = 1
            'Choice 2
            ElseIf rndExit > 50 Then
                pTunnelExit = 2
            End If
    End If
    
    'Reduce probability of entering and exiting same tunnel
    If pTunnelEnter = pTunnelExit Then
        Randomize
        rndChange = Int((100 - 1 + 1) * Rnd + 1) 'int((upperbound-lowerbound+1)*Rnd+lowerbound)
        '4 Tunnel Options
        If NumberOfTunnels = 4 Then
            'Choice 1
            If rndChange <= 10 Then
                'No Change
            'Choice 2 (3 times more likely than choice 1)
            ElseIf rndChange > 10 And rndChange <= 40 Then
                'Change
                If pTunnelExit = 1 Then
                    pTunnelExit = 2
                ElseIf pTunnelExit = 2 Then
                    pTunnelExit = 3
                ElseIf pTunnelExit = 3 Then
                    pTunnelExit = 4
                ElseIf pTunnelExit = 4 Then
                    pTunnelExit = 1
                End If
            'Choice 3 (3 times more likely than choice 1)
            ElseIf rndChange > 40 And rndChange <= 70 Then
                'Change
                If pTunnelExit = 1 Then
                    pTunnelExit = 3
                ElseIf pTunnelExit = 2 Then
                    pTunnelExit = 4
                ElseIf pTunnelExit = 3 Then
                    pTunnelExit = 1
                ElseIf pTunnelExit = 4 Then
                    pTunnelExit = 2
                End If
            'Choice 4 (3 times more likely than choice 1)
            ElseIf rndChange > 70 Then
                'Change
                If pTunnelExit = 1 Then
                    pTunnelExit = 4
                ElseIf pTunnelExit = 2 Then
                    pTunnelExit = 1
                ElseIf pTunnelExit = 3 Then
                    pTunnelExit = 2
                ElseIf pTunnelExit = 4 Then
                    pTunnelExit = 3
                End If
            End If
        '2 Tunnel Options
        ElseIf NumberOfTunnels = 2 Then
            'Choice 1
            If rndChange <= 25 Then
                'No Change
            'Choice 2 (3 times more likely than choice 1)
            Else
                'Change
                If pTunnelExit = 1 Then
                    pTunnelExit = 2
                ElseIf pTunnelExit = 2 Then
                    pTunnelExit = 1
                End If
            End If
        End If
    End If
    
    'Set first target to corresponding center corner
    If pTunnelEnter = 1 Then
        'set circular motion
        pCircularMotion = 2
        'initialize AI_Object position and motion
        AI_Object.Xcol = TunnelCoords(pTunnelEnter, 1)
        AI_Object.Yrow = TunnelCoords(pTunnelEnter, 2)
        AI_Object.nDir = 3
        AI_Object.Vx = 1
        AI_Object.Vy = 0
        'initialize first target coordinates
        pXtarget = CenterCornerCoords(3, 1)
        pYtarget = CenterCornerCoords(3, 2)
        'prevent going opposite direction
        pDirPrevent = 2
    ElseIf pTunnelEnter = 2 Then
        'set circular motion
        pCircularMotion = 1
        'initialize AI_Object position and motion
        AI_Object.Xcol = TunnelCoords(pTunnelEnter, 1)
        AI_Object.Yrow = TunnelCoords(pTunnelEnter, 2)
        AI_Object.nDir = 2
        AI_Object.Vx = -1
        AI_Object.Vy = 0
        'initialize first target coordinates
        pXtarget = CenterCornerCoords(4, 1)
        pYtarget = CenterCornerCoords(4, 2)
        'prevent going opposite direction
        pDirPrevent = 3
    ElseIf pTunnelEnter = 3 Then
        'set circular motion
        pCircularMotion = 2
        'initialize AI_Object position and motion
        AI_Object.Xcol = TunnelCoords(pTunnelEnter, 1)
        AI_Object.Yrow = TunnelCoords(pTunnelEnter, 2)
        AI_Object.nDir = 3
        AI_Object.Vx = 1
        AI_Object.Vy = 0
        'initialize first target coordinates
        pXtarget = CenterCornerCoords(3, 1)
        pYtarget = CenterCornerCoords(3, 2)
        'prevent going opposite direction
        pDirPrevent = 2
    ElseIf pTunnelEnter = 4 Then
        'set circular motion
        pCircularMotion = 1
        'initialize AI_Object position and motion
        AI_Object.Xcol = TunnelCoords(pTunnelEnter, 1)
        AI_Object.Yrow = TunnelCoords(pTunnelEnter, 2)
        AI_Object.nDir = 2
        AI_Object.Vx = -1
        AI_Object.Vy = 0
        'initialize first target coordinates
        pXtarget = CenterCornerCoords(4, 1)
        pYtarget = CenterCornerCoords(4, 2)
        'prevent going opposite direction
        pDirPrevent = 3
    End If
    
    'Start Target Counter
    pTargetCounter = 1

End Sub
