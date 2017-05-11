VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sprite_Ghosts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'=================================================================================='
'=================================== Properties ==================================='
'=================================================================================='

'AI Module
Private pAIModule As AI_Ghosts 'Holds the AI for the Ghost objects

'Ghost Name
Private pGhostName As String 'Ghost's Name for telling difference between which AI versions to follow

'Ghost Color
Private pGhostColor As Long 'Main Color of this Ghost

'Ghost Eaten
Private pEaten As Boolean 'True: Ghost has been eaten by MsP during scared state
Private pPoints As Integer 'Tells Score how much this ghost is worth
Private pPointsCounter As Integer 'Frame Timer for points animation

'Ghost Mode
Private pGhostMode As Integer '1=Attack ; 2=Scared Blue ; 3=Scared White ; 4=Eyes ; 5=Start Scared State Timer
Private pSwitchTimerMode As Integer '2=Blue mode ; 1=flashing mode
Private pSwitchTimer As Integer 'Time between color switches in Scared State (initial timer will always be TimerDuration - 5, until zero when the flashing state looses time)
Private pTimerDuration As Integer 'Length of Scared State (decreases as levels increase)
Private pMatchBoolean As Integer 'True: Match Scared State Timers on next frame
Private pScaredStateAgain As Boolean 'True: Ghost should be in scared state after eyes (used to tell ghost whether to go back into scared state after eyes if another powerdot was eaten)

'Coordinates (Top-Left Corner of C1)
Private pXcolGC As Integer 'x, or column, ghost cage coordinate
Private pYrowGC As Integer 'y, or row, ghost cage coordinate
Private pXcol As Integer 'x, or column, coordinate of the top-left-most pixel
Private pYrow As Integer 'y, or row, coordinate of the top-left-most pixel

'Velocity
Private pSpeed As Integer 'speed of Ghost
Private pVx As Integer 'x, or column, velocity of Ghost
Private pVy As Integer 'y, or row, velocity of Ghost

'Animation
Private pnDir As Integer '1=Up(200pts) ; 2=Left(400pts) ; 3=Right(800pts) ; 4=Down(1600pts)
Private pnAnim As Integer '1=Four ghost legs ; 2=Three ghost legs ; 3=Blue-4L ; 4=Blue-3L ; 5=White-4L ; 6=White-3L ; 7=Eyes Only ; 8=Points

'Current Cubes for drawing
Private pc1 As Cube 'top left cube
Private pc2 As Cube 'rop right cube
Private pc3 As Cube 'bottom left cube
Private pc4 As Cube 'bottom right cube

'Sprite Arrays (READ ONLY) (1 to 8 [anims], 1 to 4 [cubes]) each cube holds a 8x8 array
Private paUp(1 To 8, 1 To 4) As Variant 'holds all 4 cubes' pixel colors for each of the 8 frames of the Upwards animation
Private paLeft(1 To 8, 1 To 4) As Variant 'holds all 4 cubes' pixel colors for each of the 8 frames of the Leftwards animation
Private paRight(1 To 8, 1 To 4) As Variant 'holds all 4 cubes' pixel colors for each of the 8 frames of the Rightwards animation
Private paDown(1 To 8, 1 To 4) As Variant 'holds all 4 cubes' pixel colors for each of the 8 frames of the Downwards animation

'''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
' AI Module
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Property Get AIModule() As AI_Ghosts
    Set AIModule = pAIModule
End Property
Public Property Let AIModule(Value As AI_Ghosts)
    Set pAIModule = Value
End Property

'''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
' Ghost's Name
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Property Get GhostName() As String
    GhostName = pGhostName
End Property
Public Property Let GhostName(Value As String)
    pGhostName = Value
End Property

'''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
' Ghost's Main Color
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Property Get GhostColor() As Long
    GhostColor = pGhostColor
End Property
Public Property Let GhostColor(Value As Long)
    pGhostColor = Value
End Property

'''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
' Eaten Boolean
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Property Get Eaten() As Boolean
    Eaten = pEaten
End Property
Public Property Let Eaten(Value As Boolean)
    pEaten = Value
End Property

'''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
' Points
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Property Get Points() As Integer
    Points = pPoints
End Property
Public Property Let Points(Value As Integer)
    pPoints = Value
End Property

'''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
' Points Counter
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Property Get PointsCounter() As Integer
    PointsCounter = pPointsCounter
End Property
Public Property Let PointsCounter(Value As Integer)
    pPointsCounter = Value
End Property

'''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
' Ghost Mode (1=Attack ; 2=Scared Blue ; 3=Scared White ; 4=Eyes ; 5=Start Scared State Timer)
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Property Get GhostMode() As Integer
    GhostMode = pGhostMode
End Property
Public Property Let GhostMode(Value As Integer)
    pGhostMode = Value
End Property

'''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
' Switch Timer Mode
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Property Get SwitchTimerMode() As Integer
    SwitchTimerMode = pSwitchTimerMode
End Property
Public Property Let SwitchTimerMode(Value As Integer)
    pSwitchTimerMode = Value
End Property

'''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
' Switch Timer
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Property Get SwitchTimer() As Integer
    SwitchTimer = pSwitchTimer
End Property
Public Property Let SwitchTimer(Value As Integer)
    pSwitchTimer = Value
End Property

'''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
' Timer Duration
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Property Get TimerDuration() As Integer
    TimerDuration = pTimerDuration
End Property
Public Property Let TimerDuration(Value As Integer)
    pTimerDuration = Value
End Property

'''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
' Match Scared State Timers Boolean
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Property Get MatchBoolean() As Boolean
    MatchBoolean = pMatchBoolean
End Property
Public Property Let MatchBoolean(Value As Boolean)
    pMatchBoolean = Value
End Property

'''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
' Scared State Again (after eyes) Boolean
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Property Get ScaredStateAgain() As Boolean
    ScaredStateAgain = pScaredStateAgain
End Property
Public Property Let ScaredStateAgain(Value As Boolean)
    pScaredStateAgain = Value
End Property

'''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
' Ghost Cage X Coordinate (Column Number)
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Property Get XcolGC() As Integer
    XcolGC = pXcolGC
End Property
Public Property Let XcolGC(Value As Integer)
    pXcolGC = Value
End Property

'''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
' Ghost Cage Y Coordinate (Row Number)
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Property Get YrowGC() As Integer
    YrowGC = pYrowGC
End Property
Public Property Let YrowGC(Value As Integer)
    pYrowGC = Value
End Property

'''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
' X Coordinate (Column Number)
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Property Get Xcol() As Integer
    Xcol = pXcol
End Property
Public Property Let Xcol(Value As Integer)
    pXcol = Value
End Property

'''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
' Y Coordinate (Row Number)
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Property Get Yrow() As Integer
    Yrow = pYrow
End Property
Public Property Let Yrow(Value As Integer)
    pYrow = Value
End Property

'''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
' Speed
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Property Get Speed() As Integer
    Speed = pSpeed
End Property
Public Property Let Speed(Value As Integer)
    pSpeed = Value
End Property

'''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
' X Velocity
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Property Get Vx() As Integer
    Vx = pVx
End Property
Public Property Let Vx(Value As Integer)
    pVx = Value
End Property

'''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
' Y Velocity
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Property Get Vy() As Integer
    Vy = pVy
End Property
Public Property Let Vy(Value As Integer)
    pVy = Value
End Property

'''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
' Direction Number (1=Up(200pts) ; 2=Left(400pts) ; 3=Right(800pts) ; 4=Down(1600pts))
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Property Get nDir() As Integer
    nDir = pnDir
End Property
Public Property Let nDir(Value As Integer)
    pnDir = Value
End Property

'''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
' Animation Number (1=Four ghost legs ; 2=Three ghost legs ; 3=Blue-4L ; 4=Blue-3L ; 5=White-4L ; 6=White-3L ; 7=Eyes Only ; 8=Points)
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Property Get nAnim() As Integer
    nAnim = pnAnim
End Property
Public Property Let nAnim(Value As Integer)
    pnAnim = Value
End Property

'''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
' Cube 1 (top left cube)
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Property Get C1() As Cube
    Set C1 = pc1
End Property
Public Property Let C1(Value As Cube)
    Set pc1 = Value
End Property

'''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
' Cube 2
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Property Get C2() As Cube
    Set C2 = pc2
End Property
Public Property Let C2(Value As Cube)
    Set pc2 = Value
End Property

'''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
' Cube 3
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Property Get C3() As Cube
    Set C3 = pc3
End Property
Public Property Let C3(Value As Cube)
    Set pc3 = Value
End Property

'''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
' Cube 4
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Property Get C4() As Cube
    Set C4 = pc4
End Property
Public Property Let C4(Value As Cube)
    Set pc4 = Value
End Property

'''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
' Up Anim Array
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Property Get aUp(iAnimNum As Integer, iCube As Integer) As Variant
    aUp = paUp(iAnimNum, iCube)
End Property
'Public Property Let aUp(iAnimNum As Integer, iCube As Integer, Value As Variant)
'    paUp(iAnimNum, iCube) = Value
'End Property

'''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
' Left Anim Array
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Property Get aLeft(iAnimNum As Integer, iCube As Integer) As Variant
    aLeft = paLeft(iAnimNum, iCube)
End Property
'Public Property Let aLeft(iAnimNum As Integer, iCube As Integer, Value As Variant)
'    paLeft(iAnimNum, iCube) = Value
'End Property

'''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
' Right Anim Array
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Property Get aRight(iAnimNum As Integer, iCube As Integer) As Variant
    aRight = paRight(iAnimNum, iCube)
End Property
'Public Property Let aRight(iAnimNum As Integer, iCube As Integer, Value As Variant)
'    paRight(iAnimNum, iCube) = Value
'End Property

'''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
' Down Anim Array
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Property Get aDown(iAnimNum As Integer, iCube As Integer) As Variant
    aDown = paDown(iAnimNum, iCube)
End Property
'Public Property Let aDown(iAnimNum As Integer, iCube As Integer, Value As Variant)
'    paDown(iAnimNum, iCube) = Value
'End Property


'================================================================================='
'================================ Public Methods ================================='
'================================================================================='

Public Sub InitializeSprites()
'initialize all animation cubes

    'intialize Anim Arrays
    Call SetAnimLeftArray 'dir = 2, Call Left first so all others can utilize pattern already made
    Call SetAnimUpArray 'dir = 1
    Call SetAnimRightArray 'dir = 3
    Call SetAnimDownArray 'dir = 4
    Call SetFacialFeatures 'for all directions and animations

End Sub

Public Sub SetPosition()

    'Cube 1
    pc1.Xcol = pXcol
    pc1.Yrow = pYrow
    pc1.SetPixels
    'Cube 2
    pc2.Xcol = pXcol + 8
    pc2.Yrow = pYrow
    pc2.SetPixels
    'Cube 3
    pc3.Xcol = pXcol
    pc3.Yrow = pYrow + 8
    pc3.SetPixels
    'Cube 4
    pc4.Xcol = pXcol + 8
    pc4.Yrow = pYrow + 8
    pc4.SetPixels

End Sub

Public Sub UpdatePosition()

    Dim bU As Boolean 'is Up turn allowed
    Dim bL As Boolean 'is Left turn allowed
    Dim bR As Boolean 'is Right turn allowed
    Dim bD As Boolean 'is Down turn allowed
    Dim sprtRngU As Range 'sprite's potential pos after dir change up
    Dim sprtRngL As Range 'sprite's potential pos after dir change left
    Dim sprtRngR As Range 'sprite's potential pos after dir change right
    Dim sprtRngD As Range 'sprite's potential pos after dir change down
    
    'check for scared state
        If Not pEaten And pGhostMode = 5 Then 'Start the Timer from the beginning
            Me.ScaredStateTimer (True)
        ElseIf pGhostMode = 2 Or pGhostMode = 3 Then 'Continue Timer
            Me.ScaredStateTimer (False)
        End If

    'collision with MsP
        If Not pEaten And (Abs(MsP.Xcol - pXcol) <= 8) And (Abs(MsP.Yrow - pYrow) <= 8) Then
            'MsP Eaten
            If pGhostMode = 1 Then 'normal state
                'GAME OVER
                    If MsPLives = 1 And MsP.MsPMode = 1 Then
                        'stop and hide ghosts
                            'Blinky
                            Blinky.HideGhost
                            'Pinky
                            Pinky.HideGhost
                            'Inky
                            Inky.HideGhost
                            'Sue
                            Sue.HideGhost
                        'stop and hide food
                            Food.HideFood
                        'MsP loses life
                        MsPLives = MsPLives - 1 'MsPLives now = 0
                        'MsP Death Animation
                        MsP.MsPMode = 3
                'MsP Loses a Life
                    ElseIf MsP.MsPMode = 1 Then
                        'stop and hide ghosts
                            'Blinky
                            Blinky.HideGhost
                            'Pinky
                            Pinky.HideGhost
                            'Inky
                            Inky.HideGhost
                            'Sue
                            Sue.HideGhost
                        'stop and hide food
                            Food.HideFood
                        'MsP loses life
                        Set MsPLivesArr(MsPLives) = Nothing 'erase current extra life
                        MsPLives = MsPLives - 1
                        ReDim Preserve MsPLivesArr(1 To MsPLives) As Sprite_MsPacman 'resize array
                        'MsP Death Animation
                        MsP.MsPMode = 2
                    End If
            'Ghost Eaten
            ElseIf pGhostMode <> 4 And pGhostMode <> 1 Then 'scared state (eye state has no interaction)
                'Eaten Boolean
                    pEaten = True
                'Update Scores
                    ScoreUpdateBool = True
                    Call UpdateScores(Me.Points)
                'Audio
                    Call StopAudio(arrGhostEat, False) 'stop ghost eat audio
                    Call PlayAudio(arrGhostEat, False) 'play ghost eat audio
                'Ghost Mode: Eyes
                    pGhostMode = 4
                    pAIModule.MoveType = 4
                'Reverse Direction
                    Me.ReverseDirection
                'Reset Switch Timer
                    pTimerDuration = PwrDotDuration
                    pSwitchTimerMode = 2
                    pSwitchTimer = PwrDotNonFlash
                'Change Other Ghost's Points
                    If (Blinky.GhostMode <> 4) And (Blinky.GhostMode <> 1) Then: Blinky.Points = Blinky.Points * 2
                    If (Pinky.GhostMode <> 4) And (Pinky.GhostMode <> 1) Then: Pinky.Points = Pinky.Points * 2
                    If (Inky.GhostMode <> 4) And (Inky.GhostMode <> 1) Then: Inky.Points = Inky.Points * 2
                    If (Sue.GhostMode <> 4) And (Sue.GhostMode <> 1) Then: Sue.Points = Sue.Points * 2
                'Start Points Anim Coutner
                    pPointsCounter = 10 'keeps points on screen for 10 frames
            End If
        End If
    
'=============================================================================================================================================
'AI Decision Making and Movement
    'Ghost Eyes Reached Ghost Cage Entrance
    If pAIModule.MoveType = 4 And Abs(pXcol - 121) = 4 And Abs(pYrow - 109) = 0 Then
        'Enter Ghost Cage Movement Type
        pAIModule.MoveType = 0
        pAIModule.CageMoveType = 1
        pAIModule.CageMoveCounter = 1
    End If
    'Ghost Cage AI Movement
    If pAIModule.MoveType = 0 Then
        'Animate Ghost Cage Movement
            pAIModule.GhostCageAnimations Me
        'Timers
            'decrement (do not if MsP is in Death Animation Mode)
                If MsP.MsPMode = 1 Then
                    pAIModule.TimerCageDecrement
                End If
            'switch to exit ghost cage (except Blinky)
                If pGhostName <> "Blinky" And pAIModule.MoveType = 0 And pAIModule.TimerCage = 0 Then
                    'exit cage
                    pAIModule.CageMoveType = 3
                    'reset cage move counter
                    pAIModule.CageMoveCounter = 1
                End If
    'Regular AI Movement and Decision Making
    Else
        'initialize directional turn booleans
            bU = False
            bL = False
            bR = False
            bD = False
        'Change Targets
            pAIModule.TargetChange Me
        'Speed Changes
            pAIModule.ChangeSpeedCheck Me
        'Set potential new positions ("7 +" is for reaching the top-left corner of the center 4x4 square of pixels of the sprite where it touches the levelmap track)
            Set sprtRngU = Cells((7 + pYrow + (-1 * pSpeed)), (7 + pXcol + (0 * pSpeed)))
            Set sprtRngL = Cells((7 + pYrow + (0 * pSpeed)), (7 + pXcol + (-1 * pSpeed)))
            Set sprtRngR = Cells((7 + pYrow + (0 * pSpeed)), (7 + pXcol + (1 * pSpeed)))
            Set sprtRngD = Cells((7 + pYrow + (1 * pSpeed)), (7 + pXcol + (0 * pSpeed)))
        'Check for turns
            If (Not Intersect(sprtRngU, MR_Array(3, 2)) Is Nothing) Then: bU = True
            If (Not Intersect(sprtRngL, MR_Array(3, 2)) Is Nothing) Then: bL = True
            If (Not Intersect(sprtRngR, MR_Array(3, 2)) Is Nothing) Then: bR = True
            If (Not Intersect(sprtRngD, MR_Array(3, 2)) Is Nothing) Then: bD = True
        'Calculate delta coordinates
            pAIModule.CalculateSightLines Me
        'Determine directional changes if turns available
            If pAIModule.MoveType <> 0 Then
                pAIModule.DetermineDirection Me, bU, bL, bR, bD
            End If
        'Ghost Eaten in Tunnel Fix
            'Right Side Tunnels
            If pXcol < 21 Then
                'move right always
                pVx = 1
                pVy = 0
                pnDir = 3
                pAIModule.DirLast = 3
                pAIModule.DirPrevent = 2
            'Left Side Tunnels
            ElseIf pXcol > 221 Then
                'move left always
                pVx = -1
                pVy = 0
                pnDir = 2
                pAIModule.DirLast = 2
                pAIModule.DirPrevent = 3
            End If
        'Prevent movement update if points anim is visible
            If pPointsCounter > 0 Then
                pVx = 0
                pVy = 0
            End If
        'Move sprite to new coordinates
            'Half-Speed through Tunnels unless in scared state
            If (pXcol < 21 Or pXcol > 221) And pSpeed <> 2 Then
                pXcol = pXcol + (pVx * (pSpeed / 2))
                pYrow = pYrow + (pVy * (pSpeed / 2))
            'Regular Speed
            Else
                pXcol = pXcol + (pVx * pSpeed)
                pYrow = pYrow + (pVy * pSpeed)
            End If
        'Tunnel Travel
            'Left to Right
            If pXcol = 1 Then
                pXcol = 241
            'Right to Left
            ElseIf pXcol = 241 Then
                pXcol = 1
            End If
    End If
'=============================================================================================================================================
    
    'update pixel colors
        'Ghost Mode Attack
            If pGhostMode = 1 Then
                'change anim frame
                If pnAnim = 1 Then
                    pnAnim = 2
                ElseIf pnAnim = 2 Then
                    pnAnim = 1
                Else 'If moving from other ghost mode
                    pnAnim = 1
                End If
        'Ghost Mode Scared Blue
            ElseIf pGhostMode = 2 Then
                If pnAnim = 3 Then
                    pnAnim = 4
                ElseIf pnAnim = 4 Then
                    pnAnim = 3
                Else 'If moving from other ghost mode
                    pnAnim = 3
                End If
        'Ghost Mode Scared White
            ElseIf pGhostMode = 3 Then
                If pnAnim = 5 Then
                    pnAnim = 6
                ElseIf pnAnim = 6 Then
                    pnAnim = 5
                Else 'If moving from other ghost mode
                    pnAnim = 5
                End If
        'Ghost Mode Eyes
            'Points
            ElseIf pGhostMode = 4 And pPointsCounter > 0 Then
                pnAnim = 8 'points anim
                If pPoints = 200 Then: pnDir = 1 'up = 200pts
                If pPoints = 400 Then: pnDir = 2 'left = 400pts
                If pPoints = 800 Then: pnDir = 3 'right = 800pts
                If pPoints = 1600 Then: pnDir = 4 'down = 1600pts
                pPointsCounter = pPointsCounter - 1 'decrement counter
            'Eyes
            ElseIf pGhostMode = 4 And pPointsCounter = 0 Then
                'Change Anim to Eyes
                    pnAnim = 7 'only one eye animation per direction
                'Audio
                    Call StopAudio(arrPowerDotAmbient, False) 'stop Power Dot Ambient audio
                    Call PlayAudio(arrGhostEyes, False) 'play Ghost Eyes Audio
            End If
        'set anim frame
            Me.SetAnimDir
    
    'update pixel locations
        'Cube 1
        pc1.Xcol = pXcol
        pc1.Yrow = pYrow
        pc1.UpdatePixels
        'Cube 2
        pc2.Xcol = pXcol + 8
        pc2.Yrow = pYrow
        pc2.UpdatePixels
        'Cube 3
        pc3.Xcol = pXcol
        pc3.Yrow = pYrow + 8
        pc3.UpdatePixels
        'Cube 4
        pc4.Xcol = pXcol + 8
        pc4.Yrow = pYrow + 8
        pc4.UpdatePixels

End Sub

Public Sub SwitchChaseScatter()
'Switches between Chase and Scatter Phases

    'Chase to Scatter
    If (GhostPhaseType = 1) And (pAIModule.MoveType = 1) Then
        pAIModule.MoveType = 2
        Me.ReverseDirection
    'Scatter to Chase
    ElseIf (GhostPhaseType = 2) And (pAIModule.MoveType = 2) Then
        pAIModule.MoveType = 1
        Me.ReverseDirection
    End If

End Sub

Public Sub ReverseDirection()
'Reverse direction of ghost on phase changes

    'reverse direction
    If pAIModule.DirPrevent = 1 Then
        pAIModule.DirPrevent = 4
    ElseIf pAIModule.DirPrevent = 2 Then
        pAIModule.DirPrevent = 3
    ElseIf pAIModule.DirPrevent = 3 Then
        pAIModule.DirPrevent = 2
    ElseIf pAIModule.DirPrevent = 4 Then
        pAIModule.DirPrevent = 1
    End If

End Sub

Public Sub ScaredStateTimer(ByVal bStartTimer As Boolean)
'Directs the Scared State colors and timing

    'Match Scared States for Ghost Exiting Cage
    If pMatchBoolean Then
        Me.MatchScaredStateTimers
        pMatchBoolean = False
    End If
    
    'Start Timer
    If bStartTimer Then
        'start Power Dot Ambient audio if no ghosts are in eyes mode
        If Blinky.GhostMode <> 4 And Pinky.GhostMode <> 4 And Inky.GhostMode <> 4 And Sue.GhostMode <> 4 Then
            Call PlayAudio(arrPowerDotAmbient, False)
        End If
        'reset timers
        pGhostMode = 2
        pTimerDuration = PwrDotDuration
        pSwitchTimerMode = 2
        pSwitchTimer = PwrDotNonFlash
        'trigger ai scared move type
        If pAIModule.MoveType <> 0 Then
            pAIModule.MoveType = 3
        End If
        'reverse direction
        Me.ReverseDirection
    End If
    
    'reset timers
    If pTimerDuration = 0 Then
        'stop Power Dot Ambient audio
        Call StopAudio(arrPowerDotAmbient, False)
        'Ghost Mode
        pGhostMode = 1
        'reset points
        pPoints = 200
        'reset timers
        pTimerDuration = PwrDotDuration
        pSwitchTimerMode = 2
        pSwitchTimer = PwrDotNonFlash
        'change ai move type
        If pAIModule.MoveType <> 0 And pAIModule.MoveType <> 4 Then
            'change to Chase or Scatter
            pAIModule.MoveType = GhostPhaseType
        End If
    End If
    
    'blue mode
    If pSwitchTimerMode = 2 And (pGhostMode = 2) Then
        'decrement timer duration
        pTimerDuration = pTimerDuration - 1
        'switch timer ends
        If pSwitchTimer = 0 Then
            pGhostMode = 3
            pSwitchTimerMode = 1
            pSwitchTimer = 2 'length between flashes
        'decrement switch timer
        Else
            pSwitchTimer = pSwitchTimer - 1
        End If
    End If
    
    'flashing mode
    If pSwitchTimerMode = 1 And (pGhostMode = 2 Or pGhostMode = 3) Then
        'decrement timer duration
        pTimerDuration = pTimerDuration - 1
        'switch timer ends
        If pSwitchTimer = 0 Then
            'switch
            If pGhostMode = 3 Then
                pGhostMode = 2
            ElseIf pGhostMode = 2 Then
                pGhostMode = 3
            End If
            'reset switch timer
            pSwitchTimer = 2 'length between flashes
        'decrement switch timer
        Else
            pSwitchTimer = pSwitchTimer - 1
        End If
    End If

End Sub

Public Sub MatchScaredStateTimers()
'Matches a ghost exiting the cage to the point in the scared state that the rest of the ghosts are at.

    'timer duration
    If Blinky.TimerDuration <> pTimerDuration Then: pTimerDuration = Blinky.TimerDuration
    If Pinky.TimerDuration <> pTimerDuration Then: pTimerDuration = Pinky.TimerDuration
    If Inky.TimerDuration <> pTimerDuration Then: pTimerDuration = Inky.TimerDuration
    If Sue.TimerDuration <> pTimerDuration Then: pTimerDuration = Sue.TimerDuration
    
    'switch timer mode
    If Blinky.SwitchTimerMode <> pSwitchTimerMode Then: pSwitchTimerMode = Blinky.SwitchTimerMode
    If Pinky.SwitchTimerMode <> pSwitchTimerMode Then: pSwitchTimerMode = Pinky.SwitchTimerMode
    If Inky.SwitchTimerMode <> pSwitchTimerMode Then: pSwitchTimerMode = Inky.SwitchTimerMode
    If Sue.SwitchTimerMode <> pSwitchTimerMode Then: pSwitchTimerMode = Sue.SwitchTimerMode
    
    'switch timer
    If Blinky.SwitchTimer <> pSwitchTimer Then: pSwitchTimer = Blinky.SwitchTimer
    If Pinky.SwitchTimer <> pSwitchTimer Then: pSwitchTimer = Pinky.SwitchTimer
    If Inky.SwitchTimer <> pSwitchTimer Then: pSwitchTimer = Inky.SwitchTimer
    If Sue.SwitchTimer <> pSwitchTimer Then: pSwitchTimer = Sue.SwitchTimer

End Sub

Public Sub SetAnimDir()
'utilize pnDir and pnAnim to grab correct anim array and assign its contents to the pixels of C1 through C4

    Dim i As Integer 'integer iterator for 4 cubes
    Dim tC As Cube 'temporary cube object
    Dim ir As Integer 'integer iterator for pixel rows of cube
    Dim ic As Integer 'integer iterator for pixel columns of cube
    
    'Direction Up
    If pnDir = 1 Then
        'set cubes
        For ir = 1 To 8
            For ic = 1 To 8
                pc1.Pixels(ir, ic).Color = paUp(pnAnim, 1)(ir, ic)
                pc2.Pixels(ir, ic).Color = paUp(pnAnim, 2)(ir, ic)
                pc3.Pixels(ir, ic).Color = paUp(pnAnim, 3)(ir, ic)
                pc4.Pixels(ir, ic).Color = paUp(pnAnim, 4)(ir, ic)
            Next ic
        Next ir
        
    'Direction Left
    ElseIf pnDir = 2 Then
        'set cubes
        For ir = 1 To 8
            For ic = 1 To 8
                pc1.Pixels(ir, ic).Color = paLeft(pnAnim, 1)(ir, ic)
                pc2.Pixels(ir, ic).Color = paLeft(pnAnim, 2)(ir, ic)
                pc3.Pixels(ir, ic).Color = paLeft(pnAnim, 3)(ir, ic)
                pc4.Pixels(ir, ic).Color = paLeft(pnAnim, 4)(ir, ic)
            Next ic
        Next ir
        
    'Direction Right
    ElseIf pnDir = 3 Then
        'set cubes
        For ir = 1 To 8
            For ic = 1 To 8
                pc1.Pixels(ir, ic).Color = paRight(pnAnim, 1)(ir, ic)
                pc2.Pixels(ir, ic).Color = paRight(pnAnim, 2)(ir, ic)
                pc3.Pixels(ir, ic).Color = paRight(pnAnim, 3)(ir, ic)
                pc4.Pixels(ir, ic).Color = paRight(pnAnim, 4)(ir, ic)
            Next ic
        Next ir
        
    'Direction Down
    ElseIf pnDir = 4 Then
        'set cubes
        For ir = 1 To 8
            For ic = 1 To 8
                pc1.Pixels(ir, ic).Color = paDown(pnAnim, 1)(ir, ic)
                pc2.Pixels(ir, ic).Color = paDown(pnAnim, 2)(ir, ic)
                pc3.Pixels(ir, ic).Color = paDown(pnAnim, 3)(ir, ic)
                pc4.Pixels(ir, ic).Color = paDown(pnAnim, 4)(ir, ic)
            Next ic
        Next ir
        
    End If

End Sub

Public Sub Draw()
'Draws this Ghost at its current coordinates

    'Draw Cubes in order
    pc1.Draw
    pc2.Draw
    pc3.Draw
    pc4.Draw

End Sub

Public Sub HideGhost()
'Place food object back into waiting position and stop ambient audio.

    'change ai move type to "cage"
    pAIModule.MoveType = 0
    
    'reset timers
    Call ResetGhostTimers(LvlNum)
    pAIModule.TimerCageReset pGhostName
    
    'Reset Ghost Properties
    pEaten = False
    pGhostMode = 1
    pAIModule.MoveType = 0
    pAIModule.CageMoveType = 0
    pAIModule.CageMoveCounter = 0
    
    'stop movement
    pSpeed = 0
    pVx = 0
    pVy = 0
    
    'change x and y coords
    pXcol = 1
    pYrow = 253

End Sub


'=================================================================================='
'================================ Private Methods ================================='
'=================================================================================='

Private Sub Class_Initialize()

    Dim arrTemp(1 To 8, 1 To 8) As Long 'temporary array to save space for pixel colors
    Dim i1 As Integer 'integer iterator
    Dim i2 As Integer 'integer iterator
    
    'initialize AI
    Set pAIModule = New AI_Ghosts
    
    'initialize properties
    pEaten = False 'True = MsP collision during scared state
    pMatchBoolean = False 'True = Match Scared State timers on next frame
    pScaredStateAgain = False 'True = Ghost will go back into scared state after eyes state
    pPoints = 200 'first ghost eaten is always 200 pts
    pPointsCounter = 0 'points counter only necessary upon ghost eaten
    pGhostMode = 1 'starts ghost in attack mode
    pnDir = 2 'starts facing left
    pnAnim = 1 'starts with animation frame 1: Four Ghost Legs
    pSpeed = 4 'starts with 4 pixel speed
    pVx = 0 'starts with leftward velocity
    pVy = 0 'starts with zero velocity
    
    'initialize cubes
    Set pc1 = New Cube
    Set pc2 = New Cube
    Set pc3 = New Cube
    Set pc4 = New Cube
    
    'initialize arrays
    For i1 = 1 To 8
        For i2 = 1 To 4
            'upward array
            paUp(i1, i2) = arrTemp
            'leftward array
            paLeft(i1, i2) = arrTemp
            'rightward array
            paRight(i1, i2) = arrTemp
            'downward array
            paDown(i1, i2) = arrTemp
        Next i2
    Next i1

End Sub

Private Sub SetAnimLeftArray()
'Starts with background ghost color to allow copying, facial features added seperately

    Dim ir As Integer 'integer iterator for rows
    Dim ic As Integer 'integer iterator for columns
    
    'Anim: 1 (4 Ghost Legs)
        'Cube 1
            'row 2
                For ic = 7 To 8
                    paLeft(1, 1)(2, ic) = pGhostColor
                Next ic
            'row 3
                For ic = 5 To 8
                    paLeft(1, 1)(3, ic) = pGhostColor
                Next ic
            'row 4
                For ic = 4 To 8
                    paLeft(1, 1)(4, ic) = pGhostColor
                Next ic
            'rows 5-7
                For ir = 5 To 7
                    For ic = 3 To 8
                        paLeft(1, 1)(ir, ic) = pGhostColor
                    Next ic
                Next ir
            'row 8
                For ic = 2 To 8
                    paLeft(1, 1)(8, ic) = pGhostColor
                Next ic
        'Cube 2
            'y-axis reflection of Cube 1
                For ir = 1 To 8
                    For ic = 1 To 8
                        paLeft(1, 2)(ir, (9 - ic)) = paLeft(1, 1)(ir, ic)
                    Next ic
                Next ir
        'Cube 3
            'rows 1-5
                For ir = 1 To 5
                    For ic = 2 To 8
                        paLeft(1, 3)(ir, ic) = pGhostColor
                    Next ic
                Next ir
            'row 6
                paLeft(1, 3)(6, 2) = pGhostColor
                paLeft(1, 3)(6, 3) = pGhostColor
                paLeft(1, 3)(6, 5) = pGhostColor
                paLeft(1, 3)(6, 6) = pGhostColor
                paLeft(1, 3)(6, 7) = pGhostColor
            'row 7
                paLeft(1, 3)(7, 2) = pGhostColor
                paLeft(1, 3)(7, 6) = pGhostColor
                paLeft(1, 3)(7, 7) = pGhostColor
        'Cube 4
            'y-axis reflection of Cube 3
                For ir = 1 To 8
                    For ic = 1 To 8
                        paLeft(1, 4)(ir, (9 - ic)) = paLeft(1, 3)(ir, ic)
                    Next ic
                Next ir
    
    'Anim: 2 (3 Ghost Legs)
        'Cube 1
            'same as Anim 1 Cube 1
                For ir = 1 To 8
                    For ic = 1 To 8
                        paLeft(2, 1)(ir, ic) = paLeft(1, 1)(ir, ic)
                    Next ic
                Next ir
        'Cube 2
            'same as Anim 1 Cube 2
                For ir = 1 To 8
                    For ic = 1 To 8
                        paLeft(2, 2)(ir, ic) = paLeft(1, 2)(ir, ic)
                    Next ic
                Next ir
        'Cube 3
            'Top 5 rows are same as Anim 1 Cube 3
                For ir = 1 To 5
                    For ic = 1 To 8
                        paLeft(2, 3)(ir, ic) = paLeft(1, 3)(ir, ic)
                    Next ic
                Next ir
            'row 6
                paLeft(2, 3)(6, 2) = pGhostColor
                paLeft(2, 3)(6, 3) = pGhostColor
                paLeft(2, 3)(6, 4) = pGhostColor
                paLeft(2, 3)(6, 5) = pGhostColor
                paLeft(2, 3)(6, 7) = pGhostColor
                paLeft(2, 3)(6, 8) = pGhostColor
            'row 7
                paLeft(2, 3)(7, 3) = pGhostColor
                paLeft(2, 3)(7, 4) = pGhostColor
                paLeft(2, 3)(7, 8) = pGhostColor
        'Cube 4
            'y-axis reflection of Cube 3
                For ir = 1 To 8
                    For ic = 1 To 8
                        paLeft(2, 4)(ir, (9 - ic)) = paLeft(2, 3)(ir, ic)
                    Next ic
                Next ir
    
    'Anim: 3 (Scared Blue 4L)
        'Cube 1
            'copy locations of non-blacks in anim 1 and make blue
                For ir = 1 To 8
                    For ic = 1 To 8
                        If paLeft(1, 1)(ir, ic) <> Black Then
                            paLeft(3, 1)(ir, ic) = Blue
                        End If
                    Next ic
                Next ir
            'Left Eye
                For ir = 7 To 8
                    For ic = 6 To 7
                        paLeft(3, 1)(ir, ic) = Orange
                    Next ic
                Next ir
        'Cube 2
            'copy locations of non-blacks in anim 1 and make blue
                For ir = 1 To 8
                    For ic = 1 To 8
                        If paLeft(1, 2)(ir, ic) <> Black Then
                            paLeft(3, 2)(ir, ic) = Blue
                        End If
                    Next ic
                Next ir
            'Right Eye
                For ir = 7 To 8
                    For ic = 2 To 3
                        paLeft(3, 2)(ir, ic) = Orange
                    Next ic
                Next ir
        'Cube 3
            'copy locations of non-blacks in anim 1 and make blue
                For ir = 1 To 8
                    For ic = 1 To 8
                        If paLeft(1, 3)(ir, ic) <> Black Then
                            paLeft(3, 3)(ir, ic) = Blue
                        End If
                    Next ic
                Next ir
            'Mouth Left
                paLeft(3, 3)(3, 4) = Orange
                paLeft(3, 3)(3, 5) = Orange
                paLeft(3, 3)(3, 8) = Orange
                paLeft(3, 3)(4, 3) = Orange
                paLeft(3, 3)(4, 6) = Orange
                paLeft(3, 3)(4, 7) = Orange
        'Cube 4
            'reflection of Cube 3 over y-axis
                For ir = 1 To 8
                    For ic = 1 To 8
                        paLeft(3, 4)(ir, ic) = paLeft(3, 3)(ir, (9 - ic))
                    Next ic
                Next ir
    
    'Anim: 4 (Scared Blue 3L)
        'Cube 1
            'copy anim 3
                For ir = 1 To 8
                    For ic = 1 To 8
                        paLeft(4, 1)(ir, ic) = paLeft(3, 1)(ir, ic)
                    Next ic
                Next ir
        'Cube 2
            'copy anim 3
                For ir = 1 To 8
                    For ic = 1 To 8
                        paLeft(4, 2)(ir, ic) = paLeft(3, 2)(ir, ic)
                    Next ic
                Next ir
        'Cube 3
            'rows 1 to 5: copy anim 3
                For ir = 1 To 5
                    For ic = 1 To 8
                        paLeft(4, 3)(ir, ic) = paLeft(3, 3)(ir, ic)
                    Next ic
                Next ir
            'rows 6 to 8: copy locations of non-blacks in anim 2 and make blue
                For ir = 6 To 8
                    For ic = 1 To 8
                        If paLeft(2, 3)(ir, ic) <> Black Then
                            paLeft(4, 3)(ir, ic) = Blue
                        End If
                    Next ic
                Next ir
        'Cube 4
            'rows 1 to 5: copy anim 3
                For ir = 1 To 5
                    For ic = 1 To 8
                        paLeft(4, 4)(ir, ic) = paLeft(3, 4)(ir, ic)
                    Next ic
                Next ir
            'rows 6 to 8: copy locations of non-blacks in anim 2 and make blue
                For ir = 6 To 8
                    For ic = 1 To 8
                        If paLeft(2, 4)(ir, ic) <> Black Then
                            paLeft(4, 4)(ir, ic) = Blue
                        End If
                    Next ic
                Next ir
    
    'Anim: 5 (Scared White 4L)
        'Cube 1
            'copy anim 3 but make white/red
                For ir = 1 To 8
                    For ic = 1 To 8
                        If paLeft(3, 1)(ir, ic) = Blue Then
                            paLeft(5, 1)(ir, ic) = White
                        ElseIf paLeft(3, 1)(ir, ic) = Orange Then
                            paLeft(5, 1)(ir, ic) = Red
                        End If
                    Next ic
                Next ir
        'Cube 2
            'copy anim 3 but make white/red
                For ir = 1 To 8
                    For ic = 1 To 8
                        If paLeft(3, 2)(ir, ic) = Blue Then
                            paLeft(5, 2)(ir, ic) = White
                        ElseIf paLeft(3, 2)(ir, ic) = Orange Then
                            paLeft(5, 2)(ir, ic) = Red
                        End If
                    Next ic
                Next ir
        'Cube 3
            'copy anim 3 but make white/red
                For ir = 1 To 8
                    For ic = 1 To 8
                        If paLeft(3, 3)(ir, ic) = Blue Then
                            paLeft(5, 3)(ir, ic) = White
                        ElseIf paLeft(3, 3)(ir, ic) = Orange Then
                            paLeft(5, 3)(ir, ic) = Red
                        End If
                    Next ic
                Next ir
        'Cube 4
            'copy anim 3 but make white/red
                For ir = 1 To 8
                    For ic = 1 To 8
                        If paLeft(3, 4)(ir, ic) = Blue Then
                            paLeft(5, 4)(ir, ic) = White
                        ElseIf paLeft(3, 4)(ir, ic) = Orange Then
                            paLeft(5, 4)(ir, ic) = Red
                        End If
                    Next ic
                Next ir
    
    'Anim: 6 (Scared White 3L)
        'Cube 1
            'copy anim 4 but make white/red
                For ir = 1 To 8
                    For ic = 1 To 8
                        If paLeft(4, 1)(ir, ic) = Blue Then
                            paLeft(6, 1)(ir, ic) = White
                        ElseIf paLeft(4, 1)(ir, ic) = Orange Then
                            paLeft(6, 1)(ir, ic) = Red
                        End If
                    Next ic
                Next ir
        'Cube 2
            'copy anim 4 but make white/red
                For ir = 1 To 8
                    For ic = 1 To 8
                        If paLeft(4, 2)(ir, ic) = Blue Then
                            paLeft(6, 2)(ir, ic) = White
                        ElseIf paLeft(4, 2)(ir, ic) = Orange Then
                            paLeft(6, 2)(ir, ic) = Red
                        End If
                    Next ic
                Next ir
        'Cube 3
            'copy anim 4 but make white/red
                For ir = 1 To 8
                    For ic = 1 To 8
                        If paLeft(4, 3)(ir, ic) = Blue Then
                            paLeft(6, 3)(ir, ic) = White
                        ElseIf paLeft(4, 3)(ir, ic) = Orange Then
                            paLeft(6, 3)(ir, ic) = Red
                        End If
                    Next ic
                Next ir
        'Cube 4
            'copy anim 4 but make white/red
                For ir = 1 To 8
                    For ic = 1 To 8
                        If paLeft(4, 4)(ir, ic) = Blue Then
                            paLeft(6, 4)(ir, ic) = White
                        ElseIf paLeft(4, 4)(ir, ic) = Orange Then
                            paLeft(6, 4)(ir, ic) = Red
                        End If
                    Next ic
                Next ir
    
    'Anim: 7 (Eyes)
        'Cube 1
            'row 7
                For ic = 4 To 5
                    paLeft(7, 1)(7, ic) = White
                Next ic
            'row 8
                For ic = 3 To 6
                    paLeft(7, 1)(8, ic) = White
                Next ic
        'Cube 2
            'row 7
                For ic = 2 To 3
                    paLeft(7, 2)(7, ic) = White
                Next ic
            'row 8
                For ic = 1 To 4
                    paLeft(7, 2)(8, ic) = White
                Next ic
        'Cube 3
            'row 1 to 2
                'blue
                For ir = 1 To 2
                    For ic = 3 To 4
                        paLeft(7, 3)(ir, ic) = Blue
                    Next ic
                Next ir
                'white
                For ir = 1 To 2
                    For ic = 5 To 6
                        paLeft(7, 3)(ir, ic) = White
                    Next ic
                Next ir
            'row 3
                For ic = 4 To 5
                    paLeft(7, 3)(3, ic) = White
                Next ic
        'Cube 4
            'row 1 to 2
                'blue
                For ir = 1 To 2
                    For ic = 1 To 2
                        paLeft(7, 4)(ir, ic) = Blue
                    Next ic
                Next ir
                'white
                For ir = 1 To 2
                    For ic = 3 To 4
                        paLeft(7, 4)(ir, ic) = White
                    Next ic
                Next ir
            'row 3
                For ic = 2 To 3
                    paLeft(7, 4)(3, ic) = White
                Next ic
    
    'Anim: 8 (Points)
        'Cube 1
            'row 5
                paLeft(8, 1)(5, 5) = Cyan
            'row 6
                paLeft(8, 1)(6, 4) = Cyan
            'row 7
                paLeft(8, 1)(7, 3) = Cyan
            'row 8
                paLeft(8, 1)(8, 2) = Cyan
            'rows 6-8
                For ir = 6 To 8
                    paLeft(8, 1)(ir, 5) = Cyan
                    paLeft(8, 1)(ir, 8) = Cyan
                Next ir
        'Cube 2
            'row 5
                paLeft(8, 2)(5, 1) = Cyan
                paLeft(8, 2)(5, 2) = Cyan
                paLeft(8, 2)(5, 6) = Cyan
                paLeft(8, 2)(5, 7) = Cyan
            'rows 6-8
                For ir = 6 To 8
                    paLeft(8, 2)(ir, 3) = Cyan
                    paLeft(8, 2)(ir, 5) = Cyan
                    paLeft(8, 2)(ir, 8) = Cyan
                Next ir
        'Cube 3
            'row 1
                For ic = 2 To 6
                    paLeft(8, 3)(1, ic) = Cyan
                Next ic
                paLeft(8, 3)(1, 8) = Cyan
            'row 2
                paLeft(8, 3)(2, 5) = Cyan
                paLeft(8, 3)(2, 8) = Cyan
            'row 3
                paLeft(8, 3)(3, 5) = Cyan
        'Cube 4
            'rows 1-2
                For ir = 1 To 2
                    paLeft(8, 4)(ir, 3) = Cyan
                    paLeft(8, 4)(ir, 5) = Cyan
                    paLeft(8, 4)(ir, 8) = Cyan
                Next ir
            'row 3
                paLeft(8, 4)(3, 1) = Cyan
                paLeft(8, 4)(3, 2) = Cyan
                paLeft(8, 4)(3, 6) = Cyan
                paLeft(8, 4)(3, 7) = Cyan

End Sub

Private Sub SetAnimUpArray()
'Starts with background ghost color to allow copying, facial features added seperately

    Dim ir As Integer 'generic row iteration integer
    Dim ic As Integer 'generic column iteration integer
    
    'Anim: 1 (4 Ghost Legs)
        'Cube 1
            'copy ghost body from left animations
                For ir = 1 To 8
                    For ic = 1 To 8
                        paUp(1, 1)(ir, ic) = paLeft(1, 1)(ir, ic)
                    Next ic
                Next ir
            'add 2 pixels
                For ir = 6 To 7
                    paUp(1, 1)(ir, 2) = pGhostColor
                Next ir
        'Cube 2
            'copy ghost body from left animations
                For ir = 1 To 8
                    For ic = 1 To 8
                        paUp(1, 2)(ir, ic) = paLeft(1, 2)(ir, ic)
                    Next ic
                Next ir
            'add 2 pixels
                For ir = 6 To 7
                    paUp(1, 2)(ir, 7) = pGhostColor
                Next ir
        'Cube 3
            'copy ghost body from left animations
                For ir = 1 To 8
                    For ic = 1 To 8
                        paUp(1, 3)(ir, ic) = paLeft(1, 3)(ir, ic)
                    Next ic
                Next ir
        'Cube 4
            'copy ghost body from left animations
                For ir = 1 To 8
                    For ic = 1 To 8
                        paUp(1, 4)(ir, ic) = paLeft(1, 4)(ir, ic)
                    Next ic
                Next ir
    
    'Anim: 2 (3 Ghost Legs)
        'Cube 1
            'copy ghost body from left animations
                For ir = 1 To 8
                    For ic = 1 To 8
                        paUp(2, 1)(ir, ic) = paLeft(2, 1)(ir, ic)
                    Next ic
                Next ir
            'add 2 pixels
                For ir = 6 To 7
                    paUp(2, 1)(ir, 2) = pGhostColor
                Next ir
        'Cube 2
            'copy ghost body from left animations
                For ir = 1 To 8
                    For ic = 1 To 8
                        paUp(2, 2)(ir, ic) = paLeft(2, 2)(ir, ic)
                    Next ic
                Next ir
            'add 2 pixels
                For ir = 6 To 7
                    paUp(2, 2)(ir, 7) = pGhostColor
                Next ir
        'Cube 3
            'copy ghost body from left animations
                For ir = 1 To 8
                    For ic = 1 To 8
                        paUp(2, 3)(ir, ic) = paLeft(2, 3)(ir, ic)
                    Next ic
                Next ir
        'Cube 4
            'copy ghost body from left animations
                For ir = 1 To 8
                    For ic = 1 To 8
                        paUp(2, 4)(ir, ic) = paLeft(2, 4)(ir, ic)
                    Next ic
                Next ir
    
    'Anim: 3 (Scared Blue 4L)
        'Cube 1
            'copy ghost body from left animations
                For ir = 1 To 8
                    For ic = 1 To 8
                        paUp(3, 1)(ir, ic) = paLeft(3, 1)(ir, ic)
                    Next ic
                Next ir
        'Cube 2
            'copy ghost body from left animations
                For ir = 1 To 8
                    For ic = 1 To 8
                        paUp(3, 2)(ir, ic) = paLeft(3, 2)(ir, ic)
                    Next ic
                Next ir
        'Cube 3
            'copy ghost body from left animations
                For ir = 1 To 8
                    For ic = 1 To 8
                        paUp(3, 3)(ir, ic) = paLeft(3, 3)(ir, ic)
                    Next ic
                Next ir
        'Cube 4
            'copy ghost body from left animations
                For ir = 1 To 8
                    For ic = 1 To 8
                        paUp(3, 4)(ir, ic) = paLeft(3, 4)(ir, ic)
                    Next ic
                Next ir
    
    'Anim: 4 (Scared Blue 3L)
        'Cube 1
            'copy ghost body from left animations
                For ir = 1 To 8
                    For ic = 1 To 8
                        paUp(4, 1)(ir, ic) = paLeft(4, 1)(ir, ic)
                    Next ic
                Next ir
        'Cube 2
            'copy ghost body from left animations
                For ir = 1 To 8
                    For ic = 1 To 8
                        paUp(4, 2)(ir, ic) = paLeft(4, 2)(ir, ic)
                    Next ic
                Next ir
        'Cube 3
            'copy ghost body from left animations
                For ir = 1 To 8
                    For ic = 1 To 8
                        paUp(4, 3)(ir, ic) = paLeft(4, 3)(ir, ic)
                    Next ic
                Next ir
        'Cube 4
            'copy ghost body from left animations
                For ir = 1 To 8
                    For ic = 1 To 8
                        paUp(4, 4)(ir, ic) = paLeft(4, 4)(ir, ic)
                    Next ic
                Next ir
    
    'Anim: 5 (Scared White 4L)
        'Cube 1
            'copy ghost body from left animations
                For ir = 1 To 8
                    For ic = 1 To 8
                        paUp(5, 1)(ir, ic) = paLeft(5, 1)(ir, ic)
                    Next ic
                Next ir
        'Cube 2
            'copy ghost body from left animations
                For ir = 1 To 8
                    For ic = 1 To 8
                        paUp(5, 2)(ir, ic) = paLeft(5, 2)(ir, ic)
                    Next ic
                Next ir
        'Cube 3
            'copy ghost body from left animations
                For ir = 1 To 8
                    For ic = 1 To 8
                        paUp(5, 3)(ir, ic) = paLeft(5, 3)(ir, ic)
                    Next ic
                Next ir
        'Cube 4
            'copy ghost body from left animations
                For ir = 1 To 8
                    For ic = 1 To 8
                        paUp(5, 4)(ir, ic) = paLeft(5, 4)(ir, ic)
                    Next ic
                Next ir
    
    'Anim: 6 (Scared White 3L)
        'Cube 1
            'copy ghost body from left animations
                For ir = 1 To 8
                    For ic = 1 To 8
                        paUp(6, 1)(ir, ic) = paLeft(6, 1)(ir, ic)
                    Next ic
                Next ir
        'Cube 2
            'copy ghost body from left animations
                For ir = 1 To 8
                    For ic = 1 To 8
                        paUp(6, 2)(ir, ic) = paLeft(6, 2)(ir, ic)
                    Next ic
                Next ir
        'Cube 3
            'copy ghost body from left animations
                For ir = 1 To 8
                    For ic = 1 To 8
                        paUp(6, 3)(ir, ic) = paLeft(6, 3)(ir, ic)
                    Next ic
                Next ir
        'Cube 4
            'copy ghost body from left animations
                For ir = 1 To 8
                    For ic = 1 To 8
                        paUp(6, 4)(ir, ic) = paLeft(6, 4)(ir, ic)
                    Next ic
                Next ir
    
    'Anim: 7 (Eyes)
        'Cube 1
            'rows 4 to 6: White
                For ir = 4 To 6
                    For ic = 4 To 7
                        paUp(7, 1)(ir, ic) = White
                    Next ic
                Next ir
            'row 7: White
                paUp(7, 1)(7, 5) = White
                paUp(7, 1)(7, 6) = White
            'rows 3 to 4: Blue
                For ir = 3 To 4
                    For ic = 5 To 6
                        paUp(7, 1)(ir, ic) = Blue
                    Next ic
                Next ir
        'Cube 2
            'reflection of Cube 1 over y-axis
                For ir = 1 To 8
                    For ic = 1 To 8
                        paUp(7, 2)(ir, ic) = paUp(7, 1)(ir, (9 - ic))
                    Next ic
                Next ir
        'Cube 3
            'Nothing
        'Cube 4
            'Nothing
    
    'Anim: 8 (Points)
        'Cube 1
            'row 5
                For ic = 3 To 5
                    paUp(8, 1)(5, ic) = Cyan
                Next ic
            'rows 6-7
                For ir = 6 To 7
                    paUp(8, 1)(ir, 2) = Cyan
                    paUp(8, 1)(ir, 6) = Cyan
                    paUp(8, 1)(ir, 8) = Cyan
                Next ir
            'row 8
                paUp(8, 1)(8, 5) = Cyan
                paUp(8, 1)(8, 8) = Cyan
        'Cube 2
            'row 5
                paUp(8, 2)(5, 1) = Cyan
                paUp(8, 2)(5, 2) = Cyan
                paUp(8, 2)(5, 6) = Cyan
                paUp(8, 2)(5, 7) = Cyan
            'rows 6-8
                For ir = 6 To 8
                    paUp(8, 2)(ir, 3) = Cyan
                    paUp(8, 2)(ir, 5) = Cyan
                    paUp(8, 2)(ir, 8) = Cyan
                Next ir
        'Cube 3
            'row 1
                paUp(8, 3)(1, 4) = Cyan
                paUp(8, 3)(1, 8) = Cyan
            'row 2
                paUp(8, 3)(2, 3) = Cyan
                paUp(8, 3)(2, 8) = Cyan
            'row 3
                For ic = 2 To 6
                    paUp(8, 3)(3, ic) = Cyan
                Next ic
        'Cube 4
            'rows 1-2
                For ir = 1 To 2
                    paUp(8, 4)(ir, 3) = Cyan
                    paUp(8, 4)(ir, 5) = Cyan
                    paUp(8, 4)(ir, 8) = Cyan
                Next ir
            'row 3
                paUp(8, 4)(3, 1) = Cyan
                paUp(8, 4)(3, 2) = Cyan
                paUp(8, 4)(3, 6) = Cyan
                paUp(8, 4)(3, 7) = Cyan

End Sub

Private Sub SetAnimRightArray()
'Starts with background ghost color to allow copying, facial features added seperately

    Dim ir As Integer 'generic row iteration integer
    Dim ic As Integer 'generic column iteration integer
    
    'Anim: 1 (4 Ghost Legs)
        'Cube 1
            'copy ghost body from left animations
                For ir = 1 To 8
                    For ic = 1 To 8
                        paRight(1, 1)(ir, ic) = paLeft(1, 1)(ir, ic)
                    Next ic
                Next ir
        'Cube 2
            'copy ghost body from left animations
                For ir = 1 To 8
                    For ic = 1 To 8
                        paRight(1, 2)(ir, ic) = paLeft(1, 2)(ir, ic)
                    Next ic
                Next ir
        'Cube 3
            'copy ghost body from left animations
                For ir = 1 To 8
                    For ic = 1 To 8
                        paRight(1, 3)(ir, ic) = paLeft(1, 3)(ir, ic)
                    Next ic
                Next ir
        'Cube 4
            'copy ghost body from left animations
                For ir = 1 To 8
                    For ic = 1 To 8
                        paRight(1, 4)(ir, ic) = paLeft(1, 4)(ir, ic)
                    Next ic
                Next ir
    
    'Anim: 2 (3 Ghost Legs)
        'Cube 1
            'copy ghost body from left animations
                For ir = 1 To 8
                    For ic = 1 To 8
                        paRight(2, 1)(ir, ic) = paLeft(2, 1)(ir, ic)
                    Next ic
                Next ir
        'Cube 2
            'copy ghost body from left animations
                For ir = 1 To 8
                    For ic = 1 To 8
                        paRight(2, 2)(ir, ic) = paLeft(2, 2)(ir, ic)
                    Next ic
                Next ir
        'Cube 3
            'copy ghost body from left animations
                For ir = 1 To 8
                    For ic = 1 To 8
                        paRight(2, 3)(ir, ic) = paLeft(2, 3)(ir, ic)
                    Next ic
                Next ir
        'Cube 4
            'copy ghost body from left animations
                For ir = 1 To 8
                    For ic = 1 To 8
                        paRight(2, 4)(ir, ic) = paLeft(2, 4)(ir, ic)
                    Next ic
                Next ir
    
    'Anim: 3 (Scared Blue 4L)
        'Cube 1
            'copy ghost body from left animations
                For ir = 1 To 8
                    For ic = 1 To 8
                        paRight(3, 1)(ir, ic) = paLeft(3, 1)(ir, ic)
                    Next ic
                Next ir
        'Cube 2
            'copy ghost body from left animations
                For ir = 1 To 8
                    For ic = 1 To 8
                        paRight(3, 2)(ir, ic) = paLeft(3, 2)(ir, ic)
                    Next ic
                Next ir
        'Cube 3
            'copy ghost body from left animations
                For ir = 1 To 8
                    For ic = 1 To 8
                        paRight(3, 3)(ir, ic) = paLeft(3, 3)(ir, ic)
                    Next ic
                Next ir
        'Cube 4
            'copy ghost body from left animations
                For ir = 1 To 8
                    For ic = 1 To 8
                        paRight(3, 4)(ir, ic) = paLeft(3, 4)(ir, ic)
                    Next ic
                Next ir
    
    'Anim: 4 (Scared Blue 3L)
        'Cube 1
            'copy ghost body from left animations
                For ir = 1 To 8
                    For ic = 1 To 8
                        paRight(4, 1)(ir, ic) = paLeft(4, 1)(ir, ic)
                    Next ic
                Next ir
        'Cube 2
            'copy ghost body from left animations
                For ir = 1 To 8
                    For ic = 1 To 8
                        paRight(4, 2)(ir, ic) = paLeft(4, 2)(ir, ic)
                    Next ic
                Next ir
        'Cube 3
            'copy ghost body from left animations
                For ir = 1 To 8
                    For ic = 1 To 8
                        paRight(4, 3)(ir, ic) = paLeft(4, 3)(ir, ic)
                    Next ic
                Next ir
        'Cube 4
            'copy ghost body from left animations
                For ir = 1 To 8
                    For ic = 1 To 8
                        paRight(4, 4)(ir, ic) = paLeft(4, 4)(ir, ic)
                    Next ic
                Next ir
    
    'Anim: 5 (Scared White 4L)
        'Cube 1
            'copy ghost body from left animations
                For ir = 1 To 8
                    For ic = 1 To 8
                        paRight(5, 1)(ir, ic) = paLeft(5, 1)(ir, ic)
                    Next ic
                Next ir
        'Cube 2
            'copy ghost body from left animations
                For ir = 1 To 8
                    For ic = 1 To 8
                        paRight(5, 2)(ir, ic) = paLeft(5, 2)(ir, ic)
                    Next ic
                Next ir
        'Cube 3
            'copy ghost body from left animations
                For ir = 1 To 8
                    For ic = 1 To 8
                        paRight(5, 3)(ir, ic) = paLeft(5, 3)(ir, ic)
                    Next ic
                Next ir
        'Cube 4
            'copy ghost body from left animations
                For ir = 1 To 8
                    For ic = 1 To 8
                        paRight(5, 4)(ir, ic) = paLeft(5, 4)(ir, ic)
                    Next ic
                Next ir
    
    'Anim: 6 (Scared White 3L)
        'Cube 1
            'copy ghost body from left animations
                For ir = 1 To 8
                    For ic = 1 To 8
                        paRight(6, 1)(ir, ic) = paLeft(6, 1)(ir, ic)
                    Next ic
                Next ir
        'Cube 2
            'copy ghost body from left animations
                For ir = 1 To 8
                    For ic = 1 To 8
                        paRight(6, 2)(ir, ic) = paLeft(6, 2)(ir, ic)
                    Next ic
                Next ir
        'Cube 3
            'copy ghost body from left animations
                For ir = 1 To 8
                    For ic = 1 To 8
                        paRight(6, 3)(ir, ic) = paLeft(6, 3)(ir, ic)
                    Next ic
                Next ir
        'Cube 4
            'copy ghost body from left animations
                For ir = 1 To 8
                    For ic = 1 To 8
                        paRight(6, 4)(ir, ic) = paLeft(6, 4)(ir, ic)
                    Next ic
                Next ir
    
    'Anim: 7 (Eyes)
        'Cube 1
            'y-axis reflection of left anim 7 cube 2
                For ir = 1 To 8
                    For ic = 1 To 8
                        paRight(7, 1)(ir, ic) = paLeft(7, 2)(ir, (9 - ic))
                    Next ic
                Next ir
        'Cube 2
            'y-axis reflection of left anim 7 cube 1
                For ir = 1 To 8
                    For ic = 1 To 8
                        paRight(7, 2)(ir, ic) = paLeft(7, 1)(ir, (9 - ic))
                    Next ic
                Next ir
        'Cube 3
            'y-axis reflection of left anim 7 cube 4
                For ir = 1 To 8
                    For ic = 1 To 8
                        paRight(7, 3)(ir, ic) = paLeft(7, 4)(ir, (9 - ic))
                    Next ic
                Next ir
        'Cube 4
            'y-axis reflection of left anim 7 cube 3
                For ir = 1 To 8
                    For ic = 1 To 8
                        paRight(7, 4)(ir, ic) = paLeft(7, 3)(ir, (9 - ic))
                    Next ic
                Next ir
    
    'Anim: 8 (Points)
        'Cube 1
            'row 5
                For ic = 3 To 5
                    paRight(8, 1)(5, ic) = Cyan
                Next ic
            'rows 6-7
                For ir = 6 To 7
                    paRight(8, 1)(ir, 2) = Cyan
                    paRight(8, 1)(ir, 6) = Cyan
                    paRight(8, 1)(ir, 8) = Cyan
                Next ir
            'row 8
                For ic = 3 To 5
                    paRight(8, 1)(8, ic) = Cyan
                Next ic
                paRight(8, 1)(8, 8) = Cyan
            
        'Cube 2
            'row 5
                paRight(8, 2)(5, 1) = Cyan
                paRight(8, 2)(5, 2) = Cyan
                paRight(8, 2)(5, 6) = Cyan
                paRight(8, 2)(5, 7) = Cyan
            'rows 6-8
                For ir = 6 To 8
                    paRight(8, 2)(ir, 3) = Cyan
                    paRight(8, 2)(ir, 5) = Cyan
                    paRight(8, 2)(ir, 8) = Cyan
                Next ir
        'Cube 3
            'rows 1-2
                For ir = 1 To 2
                    paRight(8, 3)(ir, 2) = Cyan
                    paRight(8, 3)(ir, 6) = Cyan
                    paRight(8, 3)(ir, 8) = Cyan
                Next ir
            'row 3
                For ic = 3 To 5
                    paRight(8, 3)(3, ic) = Cyan
                Next ic
        'Cube 4
            'rows 1-2
                For ir = 1 To 2
                    paRight(8, 4)(ir, 3) = Cyan
                    paRight(8, 4)(ir, 5) = Cyan
                    paRight(8, 4)(ir, 8) = Cyan
                Next ir
            'row 3
                paRight(8, 4)(3, 1) = Cyan
                paRight(8, 4)(3, 2) = Cyan
                paRight(8, 4)(3, 6) = Cyan
                paRight(8, 4)(3, 7) = Cyan

End Sub

Private Sub SetAnimDownArray()
'Starts with background ghost color to allow copying, facial features added seperately

    Dim ir As Integer 'generic row iteration integer
    Dim ic As Integer 'generic column iteration integer
    
    'Anim: 1 (4 Ghost Legs)
        'Cube 1
            'copy ghost body from left animations
                For ir = 1 To 8
                    For ic = 1 To 8
                        paDown(1, 1)(ir, ic) = paLeft(1, 1)(ir, ic)
                    Next ic
                Next ir
        'Cube 2
            'copy ghost body from left animations
                For ir = 1 To 8
                    For ic = 1 To 8
                        paDown(1, 2)(ir, ic) = paLeft(1, 2)(ir, ic)
                    Next ic
                Next ir
        'Cube 3
            'copy ghost body from left animations
                For ir = 1 To 8
                    For ic = 1 To 8
                        paDown(1, 3)(ir, ic) = paLeft(1, 3)(ir, ic)
                    Next ic
                Next ir
        'Cube 4
            'copy ghost body from left animations
                For ir = 1 To 8
                    For ic = 1 To 8
                        paDown(1, 4)(ir, ic) = paLeft(1, 4)(ir, ic)
                    Next ic
                Next ir
    
    'Anim: 2 (3 Ghost Legs)
        'Cube 1
            'copy ghost body from left animations
                For ir = 1 To 8
                    For ic = 1 To 8
                        paDown(2, 1)(ir, ic) = paLeft(2, 1)(ir, ic)
                    Next ic
                Next ir
        'Cube 2
            'copy ghost body from left animations
                For ir = 1 To 8
                    For ic = 1 To 8
                        paDown(2, 2)(ir, ic) = paLeft(2, 2)(ir, ic)
                    Next ic
                Next ir
        'Cube 3
            'copy ghost body from left animations
                For ir = 1 To 8
                    For ic = 1 To 8
                        paDown(2, 3)(ir, ic) = paLeft(2, 3)(ir, ic)
                    Next ic
                Next ir
        'Cube 4
            'copy ghost body from left animations
                For ir = 1 To 8
                    For ic = 1 To 8
                        paDown(2, 4)(ir, ic) = paLeft(2, 4)(ir, ic)
                    Next ic
                Next ir
    
    'Anim: 3 (Scared Blue 4L)
        'Cube 1
            'copy ghost body from left animations
                For ir = 1 To 8
                    For ic = 1 To 8
                        paDown(3, 1)(ir, ic) = paLeft(3, 1)(ir, ic)
                    Next ic
                Next ir
        'Cube 2
            'copy ghost body from left animations
                For ir = 1 To 8
                    For ic = 1 To 8
                        paDown(3, 2)(ir, ic) = paLeft(3, 2)(ir, ic)
                    Next ic
                Next ir
        'Cube 3
            'copy ghost body from left animations
                For ir = 1 To 8
                    For ic = 1 To 8
                        paDown(3, 3)(ir, ic) = paLeft(3, 3)(ir, ic)
                    Next ic
                Next ir
        'Cube 4
            'copy ghost body from left animations
                For ir = 1 To 8
                    For ic = 1 To 8
                        paDown(3, 4)(ir, ic) = paLeft(3, 4)(ir, ic)
                    Next ic
                Next ir
    
    'Anim: 4 (Scared Blue 3L)
        'Cube 1
            'copy ghost body from left animations
                For ir = 1 To 8
                    For ic = 1 To 8
                        paDown(4, 1)(ir, ic) = paLeft(4, 1)(ir, ic)
                    Next ic
                Next ir
        'Cube 2
            'copy ghost body from left animations
                For ir = 1 To 8
                    For ic = 1 To 8
                        paDown(4, 2)(ir, ic) = paLeft(4, 2)(ir, ic)
                    Next ic
                Next ir
        'Cube 3
            'copy ghost body from left animations
                For ir = 1 To 8
                    For ic = 1 To 8
                        paDown(4, 3)(ir, ic) = paLeft(4, 3)(ir, ic)
                    Next ic
                Next ir
        'Cube 4
            'copy ghost body from left animations
                For ir = 1 To 8
                    For ic = 1 To 8
                        paDown(4, 4)(ir, ic) = paLeft(4, 4)(ir, ic)
                    Next ic
                Next ir
    
    'Anim: 5 (Scared White 4L)
        'Cube 1
            'copy ghost body from left animations
                For ir = 1 To 8
                    For ic = 1 To 8
                        paDown(5, 1)(ir, ic) = paLeft(5, 1)(ir, ic)
                    Next ic
                Next ir
        'Cube 2
            'copy ghost body from left animations
                For ir = 1 To 8
                    For ic = 1 To 8
                        paDown(5, 2)(ir, ic) = paLeft(5, 2)(ir, ic)
                    Next ic
                Next ir
        'Cube 3
            'copy ghost body from left animations
                For ir = 1 To 8
                    For ic = 1 To 8
                        paDown(5, 3)(ir, ic) = paLeft(5, 3)(ir, ic)
                    Next ic
                Next ir
        'Cube 4
            'copy ghost body from left animations
                For ir = 1 To 8
                    For ic = 1 To 8
                        paDown(5, 4)(ir, ic) = paLeft(5, 4)(ir, ic)
                    Next ic
                Next ir
    
    'Anim: 6 (Scared White 3L)
        'Cube 1
            'copy ghost body from left animations
                For ir = 1 To 8
                    For ic = 1 To 8
                        paDown(6, 1)(ir, ic) = paLeft(6, 1)(ir, ic)
                    Next ic
                Next ir
        'Cube 2
            'copy ghost body from left animations
                For ir = 1 To 8
                    For ic = 1 To 8
                        paDown(6, 2)(ir, ic) = paLeft(6, 2)(ir, ic)
                    Next ic
                Next ir
        'Cube 3
            'copy ghost body from left animations
                For ir = 1 To 8
                    For ic = 1 To 8
                        paDown(6, 3)(ir, ic) = paLeft(6, 3)(ir, ic)
                    Next ic
                Next ir
        'Cube 4
            'copy ghost body from left animations
                For ir = 1 To 8
                    For ic = 1 To 8
                        paDown(6, 4)(ir, ic) = paLeft(6, 4)(ir, ic)
                    Next ic
                Next ir
    
    'Anim: 7 (Eyes)
        'Cube 1
            'row 7
                For ic = 5 To 6
                    paDown(7, 1)(7, ic) = White
                Next ic
            'row 8
                For ic = 4 To 7
                    paDown(7, 1)(8, ic) = White
                Next ic
        'Cube 2
            'y-axis reflection of Cube 1
                For ir = 1 To 8
                    For ic = 1 To 8
                        paDown(7, 2)(ir, ic) = paDown(7, 1)(ir, (9 - ic))
                    Next ic
                Next ir
        'Cube 3
            'White
                For ir = 1 To 2
                    For ic = 4 To 7
                        paDown(7, 3)(ir, ic) = White
                    Next ic
                Next ir
            'Blue
                For ir = 2 To 3
                    For ic = 5 To 6
                        paDown(7, 3)(ir, ic) = Blue
                    Next ic
                Next ir
        'Cube 4
            'y-axis reflection of Cube 3
                For ir = 1 To 8
                    For ic = 1 To 8
                        paDown(7, 4)(ir, ic) = paDown(7, 3)(ir, (9 - ic))
                    Next ic
                Next ir
    
    'Anim: 8 (Points)
        'Cube 1
            'row 5
                paDown(8, 1)(5, 1) = Cyan
            'rows 5,8
                For ic = 4 To 6
                    paDown(8, 1)(5, ic) = Cyan
                    paDown(8, 1)(8, ic) = Cyan
                Next ic
            'rows 6-8
                For ir = 6 To 8
                    paDown(8, 1)(ir, 1) = Cyan
                    paDown(8, 1)(ir, 3) = Cyan
                    paDown(8, 1)(ir, 8) = Cyan
                Next ir
            
        'Cube 2
            'row 5
                paDown(8, 2)(5, 1) = Cyan
                paDown(8, 2)(5, 2) = Cyan
                paDown(8, 2)(5, 6) = Cyan
                paDown(8, 2)(5, 7) = Cyan
            'rows 6-8
                For ir = 6 To 8
                    paDown(8, 2)(ir, 3) = Cyan
                    paDown(8, 2)(ir, 5) = Cyan
                    paDown(8, 2)(ir, 8) = Cyan
                Next ir
        'Cube 3
            'rows 1-2
                For ir = 1 To 2
                    paDown(8, 3)(ir, 1) = Cyan
                    paDown(8, 3)(ir, 3) = Cyan
                    paDown(8, 3)(ir, 6) = Cyan
                    paDown(8, 3)(ir, 8) = Cyan
                Next ir
            'row 3
                paDown(8, 3)(3, 1) = Cyan
                paDown(8, 3)(3, 4) = Cyan
                paDown(8, 3)(3, 5) = Cyan
        'Cube 4
            'rows 1-2
                For ir = 1 To 2
                    paDown(8, 4)(ir, 3) = Cyan
                    paDown(8, 4)(ir, 5) = Cyan
                    paDown(8, 4)(ir, 8) = Cyan
                Next ir
            'row 3
                paDown(8, 4)(3, 1) = Cyan
                paDown(8, 4)(3, 2) = Cyan
                paDown(8, 4)(3, 6) = Cyan
                paDown(8, 4)(3, 7) = Cyan

End Sub

Private Sub SetFacialFeatures()
'Setting facial features for all animation frames in all directions.

    Dim ir As Integer 'generic row iteration integer
    Dim ic As Integer 'generic column iteration integer
    
    'Add eyes to Anims 1 and 2 from 7 for each of the 4 directions
        For ir = 1 To 8
            For ic = 1 To 8
                'Left
                    'Cube 1
                    If paLeft(7, 1)(ir, ic) <> Black Then
                        paLeft(1, 1)(ir, ic) = paLeft(7, 1)(ir, ic) 'Cube 1
                        paLeft(2, 1)(ir, ic) = paLeft(7, 1)(ir, ic) 'Cube 2
                    End If
                    'Cube 2
                    If paLeft(7, 2)(ir, ic) <> Black Then
                        paLeft(1, 2)(ir, ic) = paLeft(7, 2)(ir, ic) 'Cube 1
                        paLeft(2, 2)(ir, ic) = paLeft(7, 2)(ir, ic) 'Cube 2
                    End If
                    'Cube 3
                    If paLeft(7, 3)(ir, ic) <> Black Then
                        paLeft(1, 3)(ir, ic) = paLeft(7, 3)(ir, ic) 'Cube 1
                        paLeft(2, 3)(ir, ic) = paLeft(7, 3)(ir, ic) 'Cube 2
                    End If
                    'Cube 4
                    If paLeft(7, 4)(ir, ic) <> Black Then
                        paLeft(1, 4)(ir, ic) = paLeft(7, 4)(ir, ic) 'Cube 1
                        paLeft(2, 4)(ir, ic) = paLeft(7, 4)(ir, ic) 'Cube 2
                    End If
                'Up
                    'Cube 1
                    If paUp(7, 1)(ir, ic) <> Black Then
                        paUp(1, 1)(ir, ic) = paUp(7, 1)(ir, ic) 'Cube 1
                        paUp(2, 1)(ir, ic) = paUp(7, 1)(ir, ic) 'Cube 2
                    End If
                    'Cube 2
                    If paUp(7, 2)(ir, ic) <> Black Then
                        paUp(1, 2)(ir, ic) = paUp(7, 2)(ir, ic) 'Cube 1
                        paUp(2, 2)(ir, ic) = paUp(7, 2)(ir, ic) 'Cube 2
                    End If
                    'Cube 3
                    If paUp(7, 3)(ir, ic) <> Black Then
                        paUp(1, 3)(ir, ic) = paUp(7, 3)(ir, ic) 'Cube 1
                        paUp(2, 3)(ir, ic) = paUp(7, 3)(ir, ic) 'Cube 2
                    End If
                    'Cube 4
                    If paUp(7, 4)(ir, ic) <> Black Then
                        paUp(1, 4)(ir, ic) = paUp(7, 4)(ir, ic) 'Cube 1
                        paUp(2, 4)(ir, ic) = paUp(7, 4)(ir, ic) 'Cube 2
                    End If
                'Right
                    'Cube 1
                    If paRight(7, 1)(ir, ic) <> Black Then
                        paRight(1, 1)(ir, ic) = paRight(7, 1)(ir, ic) 'Cube 1
                        paRight(2, 1)(ir, ic) = paRight(7, 1)(ir, ic) 'Cube 2
                    End If
                    'Cube 2
                    If paRight(7, 2)(ir, ic) <> Black Then
                        paRight(1, 2)(ir, ic) = paRight(7, 2)(ir, ic) 'Cube 1
                        paRight(2, 2)(ir, ic) = paRight(7, 2)(ir, ic) 'Cube 2
                    End If
                    'Cube 3
                    If paRight(7, 3)(ir, ic) <> Black Then
                        paRight(1, 3)(ir, ic) = paRight(7, 3)(ir, ic) 'Cube 1
                        paRight(2, 3)(ir, ic) = paRight(7, 3)(ir, ic) 'Cube 2
                    End If
                    'Cube 4
                    If paRight(7, 4)(ir, ic) <> Black Then
                        paRight(1, 4)(ir, ic) = paRight(7, 4)(ir, ic) 'Cube 1
                        paRight(2, 4)(ir, ic) = paRight(7, 4)(ir, ic) 'Cube 2
                    End If
                'Down
                    'Cube 1
                    If paDown(7, 1)(ir, ic) <> Black Then
                        paDown(1, 1)(ir, ic) = paDown(7, 1)(ir, ic) 'Cube 1
                        paDown(2, 1)(ir, ic) = paDown(7, 1)(ir, ic) 'Cube 2
                    End If
                    'Cube 2
                    If paDown(7, 2)(ir, ic) <> Black Then
                        paDown(1, 2)(ir, ic) = paDown(7, 2)(ir, ic) 'Cube 1
                        paDown(2, 2)(ir, ic) = paDown(7, 2)(ir, ic) 'Cube 2
                    End If
                    'Cube 3
                    If paDown(7, 3)(ir, ic) <> Black Then
                        paDown(1, 3)(ir, ic) = paDown(7, 3)(ir, ic) 'Cube 1
                        paDown(2, 3)(ir, ic) = paDown(7, 3)(ir, ic) 'Cube 2
                    End If
                    'Cube 4
                    If paDown(7, 4)(ir, ic) <> Black Then
                        paDown(1, 4)(ir, ic) = paDown(7, 4)(ir, ic) 'Cube 1
                        paDown(2, 4)(ir, ic) = paDown(7, 4)(ir, ic) 'Cube 2
                    End If
            Next ic
        Next ir

End Sub
