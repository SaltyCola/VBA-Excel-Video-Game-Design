VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sprite_MsPacman"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'=================================================================================='
'=================================== Properties ==================================='
'=================================================================================='

'Death Animation Properties
Private pMsPMode As Integer '1=Normal ; 2=Death Animation ; 3=Game Over
Private pDeathAnimTimer As Integer 'Times the death animation

'Coordinates (Top-Left Corner of C1)
Private pXcol As Integer 'x, or column, coordinate of the top-left-most pixel
Private pYrow As Integer 'y, or row, coordinate of the top-left-most pixel

'Velocity
Private pSpeed As Integer 'speed of MsP
Private pVx As Integer 'x, or column, velocity of MsP
Private pVy As Integer 'y, or row, velocity of MsP

'Animation
Private pnDir As Integer '1=Up ; 2=Left ; 3=Right ; 4=Down
Private pnAnim As Integer '1=Open ; 2=Regular ; 3=Closed ; 4=2=Regular (corresponds to first dimension of anim arrays) (anim follows 1,2,3,4,1,2,3,4,1,etc...)

'Current Cubes for drawing
Private pc1 As Cube 'top left cube
Private pc2 As Cube 'rop right cube
Private pc3 As Cube 'bottom left cube
Private pc4 As Cube 'bottom right cube

'Sprite Arrays (READ ONLY) (1 to 4 [anims], 1 to 4 [cubes]) each cube holds a 8x8 array
Private paUp(1 To 4, 1 To 4) As Variant 'holds all 4 cubes' pixel colors for each of the 4 frames of the Upwards animation
Private paLeft(1 To 4, 1 To 4) As Variant 'holds all 4 cubes' pixel colors for each of the 4 frames of the Leftwards animation
Private paRight(1 To 4, 1 To 4) As Variant 'holds all 4 cubes' pixel colors for each of the 4 frames of the Rightwards animation
Private paDown(1 To 4, 1 To 4) As Variant 'holds all 4 cubes' pixel colors for each of the 4 frames of the Downwards animation

'''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
' Death Animation Timer
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Property Get DeathAnimTimer() As Integer
    DeathAnimTimer = pDeathAnimTimer
End Property
Public Property Let DeathAnimTimer(Value As Integer)
    pDeathAnimTimer = Value
End Property

'''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
' MsP Mode
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Property Get MsPMode() As Integer
    MsPMode = pMsPMode
End Property
Public Property Let MsPMode(Value As Integer)
    pMsPMode = Value
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
' Direction Number (1=Up ; 2=Left ; 3=Right ; 4=Down)
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
' Animation Number (1=Open ; 2=Regular ; 3=Closed ; 4=Regular)
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

    Dim sprtRngU As Range 'sprite's potential pos after dir change up
    Dim sprtRngL As Range 'sprite's potential pos after dir change left
    Dim sprtRngR As Range 'sprite's potential pos after dir change right
    Dim sprtRngD As Range 'sprite's potential pos after dir change down
    Dim sprtRngCont As Range 'sprite's potential pos after continuing current direction
    
    'MsP Updates (non-death anim)
        If pMsPMode = 1 Then
            'switch user input back on after one frame
                If BoolKeyState = False Then: BoolKeyState = True
            'Tunnel travel
                'Left to Right
                If pXcol = 1 And BoolKeyState Then
                    BoolKeyState = False 'prevent user input for one frame to prevent off-sheet crashing on left side of map
                    pXcol = 241
                'Right to Left
                ElseIf pXcol = 241 And BoolKeyState Then
                    BoolKeyState = False 'prevent user input for one frame to prevent off-sheet crashing on left side of map
                    pXcol = 1
                End If
            'set potential new positions ("7 +" is for reaching the top-left corner of the center 4x4 square of pixels of the sprite where it touches the levelmap track)
                Set sprtRngU = Cells((7 + pYrow + (-1 * pSpeed)), (7 + pXcol + (0 * pSpeed)))
                Set sprtRngL = Cells((7 + pYrow + (0 * pSpeed)), (7 + pXcol + (-1 * pSpeed)))
                Set sprtRngR = Cells((7 + pYrow + (0 * pSpeed)), (7 + pXcol + (1 * pSpeed)))
                Set sprtRngD = Cells((7 + pYrow + (1 * pSpeed)), (7 + pXcol + (0 * pSpeed)))
                Set sprtRngCont = Cells((7 + pYrow + (pVy * pSpeed)), (7 + pXcol + (pVx * pSpeed)))
            'up change
                If BoolKeyState And (KeyState(1, 2) Or KeyState(5, 2)) And (Not Intersect(sprtRngU, MR_Array(3, 2)) Is Nothing) Then
                    pnDir = 1
                    pVx = 0
                    pVy = -1
                    pXcol = pXcol + (pVx * pSpeed)
                    pYrow = pYrow + (pVy * pSpeed)
            'left change
                ElseIf BoolKeyState And (KeyState(2, 2) Or KeyState(6, 2)) And (Not Intersect(sprtRngL, MR_Array(3, 2)) Is Nothing) Then
                    pnDir = 2
                    pVx = -1
                    pVy = 0
                    pXcol = pXcol + (pVx * pSpeed)
                    pYrow = pYrow + (pVy * pSpeed)
            'right change
                ElseIf BoolKeyState And (KeyState(3, 2) Or KeyState(7, 2)) And (Not Intersect(sprtRngR, MR_Array(3, 2)) Is Nothing) Then
                    pnDir = 3
                    pVx = 1
                    pVy = 0
                    pXcol = pXcol + (pVx * pSpeed)
                    pYrow = pYrow + (pVy * pSpeed)
            'down change
                ElseIf BoolKeyState And (KeyState(4, 2) Or KeyState(8, 2)) And (Not Intersect(sprtRngD, MR_Array(3, 2)) Is Nothing) Then
                    pnDir = 4
                    pVx = 0
                    pVy = 1
                    pXcol = pXcol + (pVx * pSpeed)
                    pYrow = pYrow + (pVy * pSpeed)
            'no change (BoolKeyState not needed here)
                ElseIf Not Intersect(sprtRngCont, MR_Array(3, 2)) Is Nothing Then
                    pXcol = pXcol + (pVx * pSpeed)
                    pYrow = pYrow + (pVy * pSpeed)
            'hit wall: default to regular sprite
                Else
                    pnAnim = 2
                End If
        End If
    
    'update pixel colors
        'normal mode
            If pMsPMode = 1 Then
                'change anim frame
                If pnAnim = 1 Then
                    pnAnim = 2
                ElseIf pnAnim = 2 Then
                    pnAnim = 3
                ElseIf pnAnim = 3 Then
                    pnAnim = 4
                ElseIf pnAnim = 4 Then
                    pnAnim = 1
                End If
        'death animation
            ElseIf pMsPMode = 2 Or pMsPMode = 3 Then
                Me.DeathAnimation
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
'Draws Ms. Pacman at her current coordinates

    'Draw Cubes in order
    pc1.Draw
    pc2.Draw
    pc3.Draw
    pc4.Draw

End Sub

Public Sub DeathAnimation()

    'decrement timer
    pDeathAnimTimer = pDeathAnimTimer - 1
    
    'Frames
    If pDeathAnimTimer = 35 Then
        'prevent ambient audio repeat-timer reset
            Call AmbientAudioResetTimers_On_Off
        'stop all audio
            Call StopAudio(arrExtraLife, False)
            Call StopAudio(arrFoodAmbient, True, "A")
            Call StopAudio(arrFoodAmbient, True, "B")
            Call StopAudio(arrFoodEat, False)
            Call StopAudio(arrGameStart, False)
            Call StopAudio(arrGhostAmbient, True, "A")
            Call StopAudio(arrGhostAmbient, True, "B")
            Call StopAudio(arrGhostEat, False)
            Call StopAudio(arrGhostEyes, False)
            Call StopAudio(arrAct1, False)
            Call StopAudio(arrAct2, False)
            Call StopAudio(arrAct3, False)
            Call StopAudio(arrMenuSelect, False)
            Call StopAudio(arrMsPDeath, False)
            Call StopAudio(arrPowerDotAmbient, False)
        'STOP ALL MOVEMENT
            'MsP
                pnAnim = 2 'regular mouth anim
                pVx = 0
                pVy = 0
    ElseIf pDeathAnimTimer = 34 Then
        'STOP ALL MOVEMENT
    ElseIf pDeathAnimTimer = 33 Then
        'play MsP Death audio
            Call PlayAudio(arrMsPDeath, False)
    ElseIf pDeathAnimTimer = 32 Then
        'STOP ALL MOVEMENT
    ElseIf pDeathAnimTimer = 31 Then
        'hide ghosts
            Blinky.HideGhost
            Pinky.HideGhost
            Inky.HideGhost
            Sue.HideGhost
        'hide food
            Food.HideFood
        'change food Ambient Audio boolean
            Food.AmbientAudio = False
        'face MsP down
            pnDir = 4
    ElseIf pDeathAnimTimer = 30 Then
        'STOP ALL MOVEMENT
    ElseIf pDeathAnimTimer = 29 Then
        'STOP ALL MOVEMENT
    ElseIf pDeathAnimTimer = 28 Then
        'STOP ALL MOVEMENT
    ElseIf pDeathAnimTimer = 27 Then
        'STOP ALL MOVEMENT
    ElseIf pDeathAnimTimer = 26 Then
        'rotate
            pnDir = 2
    ElseIf pDeathAnimTimer = 25 Then
        'rotate
            pnDir = 1
    ElseIf pDeathAnimTimer = 24 Then
        'rotate
            pnDir = 3
    ElseIf pDeathAnimTimer = 23 Then
        'rotate
            pnDir = 4
    ElseIf pDeathAnimTimer = 22 Then
        'rotate
            pnDir = 2
    ElseIf pDeathAnimTimer = 21 Then
        'rotate
            pnDir = 1
    ElseIf pDeathAnimTimer = 20 Then
        'rotate
            pnDir = 3
    ElseIf pDeathAnimTimer = 19 Then
        'rotate
            pnDir = 4
    ElseIf pDeathAnimTimer = 18 Then
        'rotate
            pnDir = 2
    ElseIf pDeathAnimTimer = 17 Then
        'rotate
            pnDir = 1
    ElseIf pDeathAnimTimer = 16 Then
        'hide MsP
            pXcol = 1
            pYrow = 273
    ElseIf pDeathAnimTimer = 15 And pMsPMode <> 3 Then
        'original placements
            pnDir = 2 'facing left
            pXcol = 121
            pYrow = 205
            Blinky.nDir = 2 'facing left
            Blinky.Xcol = 121
            Blinky.Yrow = 109
            Pinky.nDir = 4 'facing down
            Pinky.Xcol = Pinky.XcolGC
            Pinky.Yrow = Pinky.YrowGC
            Inky.nDir = 1 'facing up
            Inky.Xcol = Inky.XcolGC
            Inky.Yrow = Inky.YrowGC
            Sue.nDir = 1 'facing up
            Sue.Xcol = Sue.XcolGC
            Sue.Yrow = Sue.YrowGC
        'Ready Message
            If pMsPMode <> 3 Then: txtReady.DrawText
    ElseIf pDeathAnimTimer = 14 Then
        'Ready Message
            If pMsPMode <> 3 Then: txtReady.DrawText
        'Game Over Message
            If pMsPMode = 3 Then: txtGameOver.DrawText
    ElseIf pDeathAnimTimer = 13 Then
        'Ready Message
            If pMsPMode <> 3 Then: txtReady.DrawText
        'Game Over Message
            If pMsPMode = 3 Then: txtGameOver.DrawText
    ElseIf pDeathAnimTimer = 12 Then
        'Ready Message
            If pMsPMode <> 3 Then: txtReady.DrawText
        'Game Over Message
            If pMsPMode = 3 Then: txtGameOver.DrawText
    ElseIf pDeathAnimTimer = 11 Then
        'Ready Message
            If pMsPMode <> 3 Then: txtReady.DrawText
        'Game Over Message
            If pMsPMode = 3 Then: txtGameOver.DrawText
    ElseIf pDeathAnimTimer = 10 Then
        'Ready Message
            If pMsPMode <> 3 Then: txtReady.DrawText
        'Game Over Message
            If pMsPMode = 3 Then: txtGameOver.DrawText
    ElseIf pDeathAnimTimer = 9 Then
        'Ready Message
            If pMsPMode <> 3 Then: txtReady.DrawText
        'Game Over Message
            If pMsPMode = 3 Then: txtGameOver.DrawText
    ElseIf pDeathAnimTimer = 8 Then
        'Ready Message
            If pMsPMode <> 3 Then: txtReady.DrawText
        'Game Over Message
            If pMsPMode = 3 Then: txtGameOver.DrawText
    ElseIf pDeathAnimTimer = 7 Then
        'Ready Message
            If pMsPMode <> 3 Then: txtReady.DrawText
        'Game Over Message
            If pMsPMode = 3 Then: txtGameOver.DrawText
    ElseIf pDeathAnimTimer = 6 Then
        'Ready Message
            If pMsPMode <> 3 Then: txtReady.DrawText
        'Game Over Message
            If pMsPMode = 3 Then: txtGameOver.DrawText
    ElseIf pDeathAnimTimer = 5 Then
        'Ready Message
            If pMsPMode <> 3 Then: txtReady.DrawText
        'Game Over Message
            If pMsPMode = 3 Then: txtGameOver.DrawText
    ElseIf pDeathAnimTimer = 4 Then
        'Ready Message
            If pMsPMode <> 3 Then: txtReady.DrawText
        'Game Over Message
            If pMsPMode = 3 Then: txtGameOver.DrawText
    ElseIf pDeathAnimTimer = 3 Then
        'Ready Message
            If pMsPMode <> 3 Then: txtReady.DrawText
        'Game Over Message
            If pMsPMode = 3 Then: txtGameOver.DrawText
    ElseIf pDeathAnimTimer = 2 Then
        'Ready Message
            If pMsPMode <> 3 Then: txtReady.DrawText
        'Game Over Message
            If pMsPMode = 3 Then: txtGameOver.DrawText
    ElseIf pDeathAnimTimer = 1 Then
        'Ready Message
            If pMsPMode <> 3 Then: txtReady.DrawText
        'Game Over Message
            If pMsPMode = 3 Then: txtGameOver.DrawText
    ElseIf pDeathAnimTimer = 0 And pMsPMode <> 3 Then
        'reset food timer if not already eaten this level
            If Not Food.Eaten Then: Food.AIModule.TimerReset
        'Ghost Globals
            Call ResetGhostPhaseCounter(LvlNum)
            Call ResetGhostTimers(LvlNum)
            GhostPhaseType = 2
        'reset ghost ai timers
            Blinky.AIModule.TimerCageReset "Blinky"
            Pinky.AIModule.TimerCageReset "Pinky"
            Inky.AIModule.TimerCageReset "Inky"
            Sue.AIModule.TimerCageReset "Sue"
        'Ghosts AI Move Types
            Blinky.AIModule.MoveType = 2
                Blinky.AIModule.CageMoveType = 0
                Blinky.AIModule.CageMoveCounter = 1
            Pinky.AIModule.MoveType = 0
                Pinky.AIModule.CageMoveType = 2
                Pinky.AIModule.CageMoveCounter = 1
            Inky.AIModule.MoveType = 0
                Inky.AIModule.CageMoveType = 2
                Inky.AIModule.CageMoveCounter = 1
            Sue.AIModule.MoveType = 0
                Sue.AIModule.CageMoveType = 2
                Sue.AIModule.CageMoveCounter = 1
        'reset properties
            pVx = -1
            pMsPMode = 1
            pDeathAnimTimer = 36
        'turn ambient audio repeat-timers reset back on
            Call AmbientAudioResetTimers_On_Off
    ElseIf pDeathAnimTimer = 0 And pMsPMode = 3 Then
        'Game Over Message
            If pMsPMode = 3 Then: txtGameOver.DrawText
        'End Game
            Call GameEnd
    End If

End Sub


'=================================================================================='
'================================ Private Methods ================================='
'=================================================================================='

Private Sub Class_Initialize()

    Dim arrTemp(1 To 8, 1 To 8) As Long 'temporary array to save space for pixel colors
    Dim i1 As Integer 'integer iterator
    Dim i2 As Integer 'integer iterator
    
    'initialize properties
    pMsPMode = 1 'start in normal mode
    pDeathAnimTimer = 36 '1st-decrement=1 'pause=4, down=1, pause=4, rotate(2.5x clockwise)=10, hide=1, origin=1, ready!=14, reset-properties=1
    pnDir = 2 'starts facing left
    pnAnim = 1 'starts with animation frame 1: Open
    pSpeed = 4 'starts with 4 pixel speed
    pVx = -1 'starts with leftward velocity
    pVy = 0 'starts with zero velocity
    
    'initialize cubes
    Set pc1 = New Cube
    Set pc2 = New Cube
    Set pc3 = New Cube
    Set pc4 = New Cube
    
    'initialize arrays
    For i1 = 1 To 4
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
'initializes all 4 cubes of each of the 4 leftward sprites

    Dim ir As Integer 'integer iterator for rows
    Dim ic As Integer 'integer iterator for columns
    
    'Anim: 1 (Open)

        'Cube 1
            'row 3
            paLeft(1, 1)(3, 5) = Red
            paLeft(1, 1)(3, 6) = Yellow
            paLeft(1, 1)(3, 7) = Yellow
            paLeft(1, 1)(3, 8) = Yellow
            'row 4
            paLeft(1, 1)(4, 6) = Red
            paLeft(1, 1)(4, 7) = Yellow
            paLeft(1, 1)(4, 8) = Yellow
            'row 5
            paLeft(1, 1)(5, 7) = Yellow
            paLeft(1, 1)(5, 8) = Yellow
            'row 6
            paLeft(1, 1)(6, 8) = Yellow
        
        'Cube 2
            'row 2
            paLeft(1, 2)(2, 3) = Red
            paLeft(1, 2)(2, 4) = Red
            'row 3
            paLeft(1, 2)(3, 1) = Yellow
            paLeft(1, 2)(3, 2) = Red
            paLeft(1, 2)(3, 3) = Red
            paLeft(1, 2)(3, 4) = Red
            'row 4
            paLeft(1, 2)(4, 1) = Yellow
            paLeft(1, 2)(4, 2) = Red
            paLeft(1, 2)(4, 3) = Red
            paLeft(1, 2)(4, 4) = Blue
            paLeft(1, 2)(4, 5) = Red
            'row 5
            paLeft(1, 2)(5, 1) = Black
            paLeft(1, 2)(5, 2) = Yellow
            paLeft(1, 2)(5, 3) = Yellow
            paLeft(1, 2)(5, 4) = Red
            paLeft(1, 2)(5, 5) = Blue
            paLeft(1, 2)(5, 6) = Red
            paLeft(1, 2)(5, 7) = Red
            'row 6
            paLeft(1, 2)(6, 1) = Black
            paLeft(1, 2)(6, 2) = Black
            paLeft(1, 2)(6, 3) = Yellow
            paLeft(1, 2)(6, 4) = Yellow
            paLeft(1, 2)(6, 5) = Red
            paLeft(1, 2)(6, 6) = Red
            paLeft(1, 2)(6, 7) = Red
            'row 7
            paLeft(1, 2)(7, 1) = Yellow
            paLeft(1, 2)(7, 2) = Blue
            paLeft(1, 2)(7, 3) = Black
            paLeft(1, 2)(7, 4) = Yellow
            paLeft(1, 2)(7, 5) = Red
            paLeft(1, 2)(7, 6) = Red
            'row 8
            paLeft(1, 2)(8, 2) = Yellow
            paLeft(1, 2)(8, 3) = Yellow
            paLeft(1, 2)(8, 4) = Yellow
            paLeft(1, 2)(8, 5) = Yellow
            paLeft(1, 2)(8, 6) = Yellow
        
        'Cube 3
            'row 4
            paLeft(1, 3)(4, 8) = Yellow
            'row 5
            paLeft(1, 3)(5, 7) = Yellow
            paLeft(1, 3)(5, 8) = Yellow
            'row 6
            paLeft(1, 3)(6, 6) = Red
            paLeft(1, 3)(6, 7) = Yellow
            paLeft(1, 3)(6, 8) = Yellow
            'row 7
            paLeft(1, 3)(7, 5) = Red
            paLeft(1, 3)(7, 6) = Yellow
            paLeft(1, 3)(7, 7) = Yellow
            paLeft(1, 3)(7, 8) = Yellow
        
        'Cube 4
            'row 1
            paLeft(1, 4)(1, 3) = Yellow
            paLeft(1, 4)(1, 4) = Yellow
            paLeft(1, 4)(1, 5) = Yellow
            paLeft(1, 4)(1, 6) = Yellow
            'row 2
            paLeft(1, 4)(2, 2) = Yellow
            paLeft(1, 4)(2, 3) = Yellow
            paLeft(1, 4)(2, 4) = Yellow
            paLeft(1, 4)(2, 5) = Yellow
            paLeft(1, 4)(2, 6) = Yellow
            'row 3
            paLeft(1, 4)(3, 1) = Yellow
            paLeft(1, 4)(3, 2) = Yellow
            paLeft(1, 4)(3, 3) = Black
            paLeft(1, 4)(3, 4) = Yellow
            paLeft(1, 4)(3, 5) = Yellow
            paLeft(1, 4)(3, 6) = Yellow
            'row 4
            paLeft(1, 4)(4, 1) = Yellow
            paLeft(1, 4)(4, 2) = Yellow
            paLeft(1, 4)(4, 3) = Yellow
            paLeft(1, 4)(4, 4) = Yellow
            paLeft(1, 4)(4, 5) = Yellow
            'row 5
            paLeft(1, 4)(5, 1) = Yellow
            paLeft(1, 4)(5, 2) = Yellow
            paLeft(1, 4)(5, 3) = Yellow
            paLeft(1, 4)(5, 4) = Yellow
            paLeft(1, 4)(5, 5) = Yellow
            'row 6
            paLeft(1, 4)(6, 1) = Yellow
            paLeft(1, 4)(6, 2) = Yellow
            paLeft(1, 4)(6, 3) = Yellow
            paLeft(1, 4)(6, 4) = Yellow
            'row 7
            paLeft(1, 4)(7, 1) = Yellow
            paLeft(1, 4)(7, 2) = Yellow
    
    'Anim: 2 (Regular)
        
        'Cube 1
            'row 3
            paLeft(2, 1)(3, 6) = Yellow
            paLeft(2, 1)(3, 7) = Yellow
            paLeft(2, 1)(3, 8) = Yellow
            'row 4
            paLeft(2, 1)(4, 4) = Yellow
            paLeft(2, 1)(4, 5) = Yellow
            paLeft(2, 1)(4, 6) = Yellow
            paLeft(2, 1)(4, 7) = Yellow
            paLeft(2, 1)(4, 8) = Yellow
            'row 5
            paLeft(2, 1)(5, 3) = Yellow
            paLeft(2, 1)(5, 4) = Yellow
            paLeft(2, 1)(5, 5) = Yellow
            paLeft(2, 1)(5, 6) = Yellow
            paLeft(2, 1)(5, 7) = Yellow
            paLeft(2, 1)(5, 8) = Yellow
            'row 6
            paLeft(2, 1)(6, 3) = Red
            paLeft(2, 1)(6, 4) = Red
            paLeft(2, 1)(6, 5) = Yellow
            paLeft(2, 1)(6, 6) = Yellow
            paLeft(2, 1)(6, 7) = Yellow
            paLeft(2, 1)(6, 8) = Black
            'row 7
            paLeft(2, 1)(7, 7) = Yellow
            paLeft(2, 1)(7, 8) = Yellow
        
        'Cube 2
            'row 2
            paLeft(2, 2)(2, 3) = Red
            paLeft(2, 2)(2, 4) = Red
            'row 3
            paLeft(2, 2)(3, 1) = Yellow
            paLeft(2, 2)(3, 2) = Red
            paLeft(2, 2)(3, 3) = Red
            paLeft(2, 2)(3, 4) = Red
            'row 4
            paLeft(2, 2)(4, 1) = Yellow
            paLeft(2, 2)(4, 2) = Red
            paLeft(2, 2)(4, 3) = Red
            paLeft(2, 2)(4, 4) = Blue
            paLeft(2, 2)(4, 5) = Red
            'row 5
            paLeft(2, 2)(5, 1) = Yellow
            paLeft(2, 2)(5, 2) = Yellow
            paLeft(2, 2)(5, 3) = Yellow
            paLeft(2, 2)(5, 4) = Red
            paLeft(2, 2)(5, 5) = Blue
            paLeft(2, 2)(5, 6) = Red
            paLeft(2, 2)(5, 7) = Red
            'row 6
            paLeft(2, 2)(6, 1) = Black
            paLeft(2, 2)(6, 2) = Yellow
            paLeft(2, 2)(6, 3) = Yellow
            paLeft(2, 2)(6, 4) = Yellow
            paLeft(2, 2)(6, 5) = Red
            paLeft(2, 2)(6, 6) = Red
            paLeft(2, 2)(6, 7) = Red
            'row 7
            paLeft(2, 2)(7, 1) = Blue
            paLeft(2, 2)(7, 2) = Black
            paLeft(2, 2)(7, 3) = Yellow
            paLeft(2, 2)(7, 4) = Yellow
            paLeft(2, 2)(7, 5) = Red
            paLeft(2, 2)(7, 6) = Red
            'row 8
            paLeft(2, 2)(8, 1) = Yellow
            paLeft(2, 2)(8, 2) = Yellow
            paLeft(2, 2)(8, 3) = Yellow
            paLeft(2, 2)(8, 4) = Yellow
            paLeft(2, 2)(8, 5) = Yellow
            paLeft(2, 2)(8, 6) = Yellow
        
        'Cube 3
            'row 3
            paLeft(2, 3)(3, 7) = Yellow
            paLeft(2, 3)(3, 8) = Yellow
            'row 4
            paLeft(2, 3)(4, 3) = Red
            paLeft(2, 3)(4, 4) = Red
            paLeft(2, 3)(4, 5) = Yellow
            paLeft(2, 3)(4, 6) = Yellow
            paLeft(2, 3)(4, 7) = Yellow
            paLeft(2, 3)(4, 8) = Yellow
            'row 5
            paLeft(2, 3)(5, 3) = Yellow
            paLeft(2, 3)(5, 4) = Yellow
            paLeft(2, 3)(5, 5) = Yellow
            paLeft(2, 3)(5, 6) = Yellow
            paLeft(2, 3)(5, 7) = Yellow
            paLeft(2, 3)(5, 8) = Yellow
            'row 6
            paLeft(2, 3)(6, 4) = Yellow
            paLeft(2, 3)(6, 5) = Yellow
            paLeft(2, 3)(6, 6) = Yellow
            paLeft(2, 3)(6, 7) = Yellow
            paLeft(2, 3)(6, 8) = Yellow
            'row 7
            paLeft(2, 3)(7, 6) = Yellow
            paLeft(2, 3)(7, 7) = Yellow
            paLeft(2, 3)(7, 8) = Yellow
        
        'Cube 4
            'row 1
            paLeft(2, 4)(1, 3) = Yellow
            paLeft(2, 4)(1, 4) = Yellow
            paLeft(2, 4)(1, 5) = Yellow
            paLeft(2, 4)(1, 6) = Yellow
            'row 2
            paLeft(2, 4)(2, 1) = Yellow
            paLeft(2, 4)(2, 2) = Yellow
            paLeft(2, 4)(2, 3) = Yellow
            paLeft(2, 4)(2, 4) = Yellow
            paLeft(2, 4)(2, 5) = Yellow
            paLeft(2, 4)(2, 6) = Yellow
            'row 3
            paLeft(2, 4)(3, 1) = Yellow
            paLeft(2, 4)(3, 2) = Yellow
            paLeft(2, 4)(3, 3) = Black
            paLeft(2, 4)(3, 4) = Yellow
            paLeft(2, 4)(3, 5) = Yellow
            paLeft(2, 4)(3, 6) = Yellow
            'row 4
            paLeft(2, 4)(4, 1) = Yellow
            paLeft(2, 4)(4, 2) = Yellow
            paLeft(2, 4)(4, 3) = Yellow
            paLeft(2, 4)(4, 4) = Yellow
            paLeft(2, 4)(4, 5) = Yellow
            'row 5
            paLeft(2, 4)(5, 1) = Yellow
            paLeft(2, 4)(5, 2) = Yellow
            paLeft(2, 4)(5, 3) = Yellow
            paLeft(2, 4)(5, 4) = Yellow
            paLeft(2, 4)(5, 5) = Yellow
            'row 6
            paLeft(2, 4)(6, 1) = Yellow
            paLeft(2, 4)(6, 2) = Yellow
            paLeft(2, 4)(6, 3) = Yellow
            paLeft(2, 4)(6, 4) = Yellow
            'row 7
            paLeft(2, 4)(7, 1) = Yellow
            paLeft(2, 4)(7, 2) = Yellow
    
    'Anim: 3 (Closed)
        
        'Cube 1
            'row 3
            paLeft(3, 1)(3, 6) = Yellow
            paLeft(3, 1)(3, 7) = Yellow
            paLeft(3, 1)(3, 8) = Yellow
            'row 4
            paLeft(3, 1)(4, 4) = Yellow
            paLeft(3, 1)(4, 5) = Yellow
            paLeft(3, 1)(4, 6) = Yellow
            paLeft(3, 1)(4, 7) = Yellow
            paLeft(3, 1)(4, 8) = Yellow
            'row 5
            paLeft(3, 1)(5, 3) = Yellow
            paLeft(3, 1)(5, 4) = Yellow
            paLeft(3, 1)(5, 5) = Yellow
            paLeft(3, 1)(5, 6) = Yellow
            paLeft(3, 1)(5, 7) = Yellow
            paLeft(3, 1)(5, 8) = Yellow
            'row 6
            paLeft(3, 1)(6, 3) = Yellow
            paLeft(3, 1)(6, 4) = Yellow
            paLeft(3, 1)(6, 5) = Yellow
            paLeft(3, 1)(6, 6) = Yellow
            paLeft(3, 1)(6, 7) = Yellow
            paLeft(3, 1)(6, 8) = Yellow
            'row 7
            paLeft(3, 1)(7, 2) = Yellow
            paLeft(3, 1)(7, 3) = Yellow
            paLeft(3, 1)(7, 4) = Yellow
            paLeft(3, 1)(7, 5) = Yellow
            paLeft(3, 1)(7, 6) = Yellow
            paLeft(3, 1)(7, 7) = Yellow
            paLeft(3, 1)(7, 8) = Black
            'row 8
            paLeft(3, 1)(8, 1) = Red
            paLeft(3, 1)(8, 2) = Yellow
            paLeft(3, 1)(8, 3) = Yellow
            paLeft(3, 1)(8, 4) = Yellow
            paLeft(3, 1)(8, 5) = Yellow
            paLeft(3, 1)(8, 6) = Yellow
            paLeft(3, 1)(8, 7) = Yellow
            paLeft(3, 1)(8, 8) = Yellow
        
        'Cube 2
            'row 2
            paLeft(3, 2)(2, 3) = Red
            paLeft(3, 2)(2, 4) = Red
            'row 3
            paLeft(3, 2)(3, 1) = Yellow
            paLeft(3, 2)(3, 2) = Red
            paLeft(3, 2)(3, 3) = Red
            paLeft(3, 2)(3, 4) = Red
            'row 4
            paLeft(3, 2)(4, 1) = Yellow
            paLeft(3, 2)(4, 2) = Red
            paLeft(3, 2)(4, 3) = Red
            paLeft(3, 2)(4, 4) = Blue
            paLeft(3, 2)(4, 5) = Red
            'row 5
            paLeft(3, 2)(5, 1) = Yellow
            paLeft(3, 2)(5, 2) = Yellow
            paLeft(3, 2)(5, 3) = Yellow
            paLeft(3, 2)(5, 4) = Red
            paLeft(3, 2)(5, 5) = Blue
            paLeft(3, 2)(5, 6) = Red
            paLeft(3, 2)(5, 7) = Red
            'row 6
            paLeft(3, 2)(6, 1) = Yellow
            paLeft(3, 2)(6, 2) = Yellow
            paLeft(3, 2)(6, 3) = Yellow
            paLeft(3, 2)(6, 4) = Yellow
            paLeft(3, 2)(6, 5) = Red
            paLeft(3, 2)(6, 6) = Red
            paLeft(3, 2)(6, 7) = Red
            'row 7
            paLeft(3, 2)(7, 1) = Black
            paLeft(3, 2)(7, 2) = Black
            paLeft(3, 2)(7, 3) = Black
            paLeft(3, 2)(7, 4) = Yellow
            paLeft(3, 2)(7, 5) = Red
            paLeft(3, 2)(7, 6) = Red
            'row 8
            paLeft(3, 2)(8, 1) = Yellow
            paLeft(3, 2)(8, 2) = Yellow
            paLeft(3, 2)(8, 3) = Yellow
            paLeft(3, 2)(8, 4) = Yellow
            paLeft(3, 2)(8, 5) = Yellow
            paLeft(3, 2)(8, 6) = Yellow
        
        'Cube 3
            'row 1
            paLeft(3, 3)(1, 1) = Red
            paLeft(3, 3)(1, 2) = Red
            paLeft(3, 3)(1, 3) = Yellow
            paLeft(3, 3)(1, 4) = Yellow
            paLeft(3, 3)(1, 5) = Yellow
            paLeft(3, 3)(1, 6) = Yellow
            paLeft(3, 3)(1, 7) = Yellow
            paLeft(3, 3)(1, 8) = Yellow
            'row 2
            paLeft(3, 3)(2, 1) = Red
            paLeft(3, 3)(2, 2) = Yellow
            paLeft(3, 3)(2, 3) = Yellow
            paLeft(3, 3)(2, 4) = Yellow
            paLeft(3, 3)(2, 5) = Yellow
            paLeft(3, 3)(2, 6) = Yellow
            paLeft(3, 3)(2, 7) = Yellow
            paLeft(3, 3)(2, 8) = Yellow
            'row 3
            paLeft(3, 3)(3, 2) = Yellow
            paLeft(3, 3)(3, 3) = Yellow
            paLeft(3, 3)(3, 4) = Yellow
            paLeft(3, 3)(3, 5) = Yellow
            paLeft(3, 3)(3, 6) = Yellow
            paLeft(3, 3)(3, 7) = Yellow
            paLeft(3, 3)(3, 8) = Yellow
            'row 4
            paLeft(3, 3)(4, 3) = Yellow
            paLeft(3, 3)(4, 4) = Yellow
            paLeft(3, 3)(4, 5) = Yellow
            paLeft(3, 3)(4, 6) = Yellow
            paLeft(3, 3)(4, 7) = Yellow
            paLeft(3, 3)(4, 8) = Yellow
            'row 5
            paLeft(3, 3)(5, 3) = Yellow
            paLeft(3, 3)(5, 4) = Yellow
            paLeft(3, 3)(5, 5) = Yellow
            paLeft(3, 3)(5, 6) = Yellow
            paLeft(3, 3)(5, 7) = Yellow
            paLeft(3, 3)(5, 8) = Yellow
            'row 6
            paLeft(3, 3)(6, 4) = Yellow
            paLeft(3, 3)(6, 5) = Yellow
            paLeft(3, 3)(6, 6) = Yellow
            paLeft(3, 3)(6, 7) = Yellow
            paLeft(3, 3)(6, 8) = Yellow
            'row 7
            paLeft(3, 3)(7, 6) = Yellow
            paLeft(3, 3)(7, 7) = Yellow
            paLeft(3, 3)(7, 8) = Yellow
        
        'Cube 4
            'row 1
            paLeft(3, 4)(1, 1) = Yellow
            paLeft(3, 4)(1, 2) = Yellow
            paLeft(3, 4)(1, 3) = Yellow
            paLeft(3, 4)(1, 4) = Yellow
            paLeft(3, 4)(1, 5) = Yellow
            paLeft(3, 4)(1, 6) = Yellow
            'row 2
            paLeft(3, 4)(2, 1) = Yellow
            paLeft(3, 4)(2, 2) = Yellow
            paLeft(3, 4)(2, 3) = Yellow
            paLeft(3, 4)(2, 4) = Yellow
            paLeft(3, 4)(2, 5) = Yellow
            paLeft(3, 4)(2, 6) = Yellow
            'row 3
            paLeft(3, 4)(3, 1) = Yellow
            paLeft(3, 4)(3, 2) = Yellow
            paLeft(3, 4)(3, 3) = Black
            paLeft(3, 4)(3, 4) = Yellow
            paLeft(3, 4)(3, 5) = Yellow
            paLeft(3, 4)(3, 6) = Yellow
            'row 4
            paLeft(3, 4)(4, 1) = Yellow
            paLeft(3, 4)(4, 2) = Yellow
            paLeft(3, 4)(4, 3) = Yellow
            paLeft(3, 4)(4, 4) = Yellow
            paLeft(3, 4)(4, 5) = Yellow
            'row 5
            paLeft(3, 4)(5, 1) = Yellow
            paLeft(3, 4)(5, 2) = Yellow
            paLeft(3, 4)(5, 3) = Yellow
            paLeft(3, 4)(5, 4) = Yellow
            paLeft(3, 4)(5, 5) = Yellow
            'row 6
            paLeft(3, 4)(6, 1) = Yellow
            paLeft(3, 4)(6, 2) = Yellow
            paLeft(3, 4)(6, 3) = Yellow
            paLeft(3, 4)(6, 4) = Yellow
            'row 7
            paLeft(3, 4)(7, 1) = Yellow
            paLeft(3, 4)(7, 2) = Yellow
    
    'Anim: 4 (Regular (same as 2))
    
        For ir = 1 To 8
            For ic = 1 To 8
                'Cube 1
                paLeft(4, 1)(ir, ic) = paLeft(2, 1)(ir, ic)
                'Cube 2
                paLeft(4, 2)(ir, ic) = paLeft(2, 2)(ir, ic)
                'Cube 3
                paLeft(4, 3)(ir, ic) = paLeft(2, 3)(ir, ic)
                'Cube 4
                paLeft(4, 4)(ir, ic) = paLeft(2, 4)(ir, ic)
            Next ic
        Next ir

End Sub

Private Sub SetAnimUpArray()
'initializes all 4 cubes of each of the 4 upward sprites

    Dim ir As Integer 'generic row iteration integer
    Dim ic As Integer 'generic column iteration integer
    
    'Anim: 1 (Open)
    
        'Cube 1
            For ir = 1 To 8
                For ic = 1 To 8
                    paUp(1, 1)(ir, ic) = paLeft(1, 3)((9 - ic), ir)
                Next ic
            Next ir
        
        'Cube 2
            For ir = 1 To 8
                For ic = 1 To 8
                    paUp(1, 2)(ir, ic) = paLeft(1, 1)((9 - ic), ir)
                Next ic
            Next ir
        
        'Cube 3
            For ir = 1 To 8
                For ic = 1 To 8
                    paUp(1, 3)(ir, ic) = paLeft(1, 4)((9 - ic), ir)
                Next ic
            Next ir
        
        'Cube 4
            For ir = 1 To 8
                For ic = 1 To 8
                    paUp(1, 4)(ir, ic) = paLeft(1, 2)((9 - ic), ir)
                Next ic
            Next ir
    
    'Anim: 2 (Regular)
    
        'Cube 1
            For ir = 1 To 8
                For ic = 1 To 8
                    paUp(2, 1)(ir, ic) = paLeft(2, 3)((9 - ic), ir)
                Next ic
            Next ir
        
        'Cube 2
            For ir = 1 To 8
                For ic = 1 To 8
                    paUp(2, 2)(ir, ic) = paLeft(2, 1)((9 - ic), ir)
                Next ic
            Next ir
        
        'Cube 3
            For ir = 1 To 8
                For ic = 1 To 8
                    paUp(2, 3)(ir, ic) = paLeft(2, 4)((9 - ic), ir)
                Next ic
            Next ir
        
        'Cube 4
            For ir = 1 To 8
                For ic = 1 To 8
                    paUp(2, 4)(ir, ic) = paLeft(2, 2)((9 - ic), ir)
                Next ic
            Next ir
    
    'Anim: 3 (Closed)
    
        'Cube 1
            For ir = 1 To 8
                For ic = 1 To 8
                    paUp(3, 1)(ir, ic) = paLeft(3, 3)((9 - ic), ir)
                Next ic
            Next ir
        
        'Cube 2
            For ir = 1 To 8
                For ic = 1 To 8
                    paUp(3, 2)(ir, ic) = paLeft(3, 1)((9 - ic), ir)
                Next ic
            Next ir
        
        'Cube 3
            For ir = 1 To 8
                For ic = 1 To 8
                    paUp(3, 3)(ir, ic) = paLeft(3, 4)((9 - ic), ir)
                Next ic
            Next ir
        
        'Cube 4
            For ir = 1 To 8
                For ic = 1 To 8
                    paUp(3, 4)(ir, ic) = paLeft(3, 2)((9 - ic), ir)
                Next ic
            Next ir
    
    'Anim: 4 (Regular (same as 2))
    
        For ir = 1 To 8
            For ic = 1 To 8
                'Cube 1
                paUp(4, 1)(ir, ic) = paUp(2, 1)(ir, ic)
                'Cube 2
                paUp(4, 2)(ir, ic) = paUp(2, 2)(ir, ic)
                'Cube 3
                paUp(4, 3)(ir, ic) = paUp(2, 3)(ir, ic)
                'Cube 4
                paUp(4, 4)(ir, ic) = paUp(2, 4)(ir, ic)
            Next ic
        Next ir

End Sub

Private Sub SetAnimRightArray()
'initializes all 4 cubes of each of the 4 rightward sprites

    Dim ir As Integer 'generic row iteration integer
    Dim ic As Integer 'generic column iteration integer
    
    'Anim: 1 (Open)
    
        'Cube 1
            For ir = 1 To 8
                For ic = 1 To 8
                    paRight(1, 1)(ir, ic) = paLeft(1, 2)(ir, (9 - ic))
                Next ic
            Next ir
        
        'Cube 2
            For ir = 1 To 8
                For ic = 1 To 8
                    paRight(1, 2)(ir, ic) = paLeft(1, 1)(ir, (9 - ic))
                Next ic
            Next ir
        
        'Cube 3
            For ir = 1 To 8
                For ic = 1 To 8
                    paRight(1, 3)(ir, ic) = paLeft(1, 4)(ir, (9 - ic))
                Next ic
            Next ir
        
        'Cube 4
            For ir = 1 To 8
                For ic = 1 To 8
                    paRight(1, 4)(ir, ic) = paLeft(1, 3)(ir, (9 - ic))
                Next ic
            Next ir
    
    'Anim: 2 (Regular)
    
        'Cube 1
            For ir = 1 To 8
                For ic = 1 To 8
                    paRight(2, 1)(ir, ic) = paLeft(2, 2)(ir, (9 - ic))
                Next ic
            Next ir
        
        'Cube 2
            For ir = 1 To 8
                For ic = 1 To 8
                    paRight(2, 2)(ir, ic) = paLeft(2, 1)(ir, (9 - ic))
                Next ic
            Next ir
        
        'Cube 3
            For ir = 1 To 8
                For ic = 1 To 8
                    paRight(2, 3)(ir, ic) = paLeft(2, 4)(ir, (9 - ic))
                Next ic
            Next ir
        
        'Cube 4
            For ir = 1 To 8
                For ic = 1 To 8
                    paRight(2, 4)(ir, ic) = paLeft(2, 3)(ir, (9 - ic))
                Next ic
            Next ir
    
    'Anim: 3 (Closed)
    
        'Cube 1
            For ir = 1 To 8
                For ic = 1 To 8
                    paRight(3, 1)(ir, ic) = paLeft(3, 2)(ir, (9 - ic))
                Next ic
            Next ir
        
        'Cube 2
            For ir = 1 To 8
                For ic = 1 To 8
                    paRight(3, 2)(ir, ic) = paLeft(3, 1)(ir, (9 - ic))
                Next ic
            Next ir
        
        'Cube 3
            For ir = 1 To 8
                For ic = 1 To 8
                    paRight(3, 3)(ir, ic) = paLeft(3, 4)(ir, (9 - ic))
                Next ic
            Next ir
        
        'Cube 4
            For ir = 1 To 8
                For ic = 1 To 8
                    paRight(3, 4)(ir, ic) = paLeft(3, 3)(ir, (9 - ic))
                Next ic
            Next ir
    
    'Anim: 4 (Regular (same as 2))
    
        For ir = 1 To 8
            For ic = 1 To 8
                'Cube 1
                paRight(4, 1)(ir, ic) = paRight(2, 1)(ir, ic)
                'Cube 2
                paRight(4, 2)(ir, ic) = paRight(2, 2)(ir, ic)
                'Cube 3
                paRight(4, 3)(ir, ic) = paRight(2, 3)(ir, ic)
                'Cube 4
                paRight(4, 4)(ir, ic) = paRight(2, 4)(ir, ic)
            Next ic
        Next ir

End Sub

Private Sub SetAnimDownArray()
'initializes all 4 cubes of each of the 4 downward sprites

    Dim ir As Integer 'generic row iteration integer
    Dim ic As Integer 'generic column iteration integer
    
    'Anim: 1 (Open)
    
        'Cube 1
            For ir = 1 To 8
                For ic = 1 To 8
                    paDown(1, 1)(ir, ic) = paLeft(1, 2)(ic, (9 - ir))
                Next ic
            Next ir
        
        'Cube 2
            For ir = 1 To 8
                For ic = 1 To 8
                    paDown(1, 2)(ir, ic) = paLeft(1, 4)(ic, (9 - ir))
                Next ic
            Next ir
        
        'Cube 3
            For ir = 1 To 8
                For ic = 1 To 8
                    paDown(1, 3)(ir, ic) = paLeft(1, 1)(ic, (9 - ir))
                Next ic
            Next ir
        
        'Cube 4
            For ir = 1 To 8
                For ic = 1 To 8
                    paDown(1, 4)(ir, ic) = paLeft(1, 3)(ic, (9 - ir))
                Next ic
            Next ir
    
    'Anim: 2 (Regular)
    
        'Cube 1
            For ir = 1 To 8
                For ic = 1 To 8
                    paDown(2, 1)(ir, ic) = paLeft(2, 2)(ic, (9 - ir))
                Next ic
            Next ir
        
        'Cube 2
            For ir = 1 To 8
                For ic = 1 To 8
                    paDown(2, 2)(ir, ic) = paLeft(2, 4)(ic, (9 - ir))
                Next ic
            Next ir
        
        'Cube 3
            For ir = 1 To 8
                For ic = 1 To 8
                    paDown(2, 3)(ir, ic) = paLeft(2, 1)(ic, (9 - ir))
                Next ic
            Next ir
        
        'Cube 4
            For ir = 1 To 8
                For ic = 1 To 8
                    paDown(2, 4)(ir, ic) = paLeft(2, 3)(ic, (9 - ir))
                Next ic
            Next ir
    
    'Anim: 3 (Closed)
    
        'Cube 1
            For ir = 1 To 8
                For ic = 1 To 8
                    paDown(3, 1)(ir, ic) = paLeft(3, 2)(ic, (9 - ir))
                Next ic
            Next ir
        
        'Cube 2
            For ir = 1 To 8
                For ic = 1 To 8
                    paDown(3, 2)(ir, ic) = paLeft(3, 4)(ic, (9 - ir))
                Next ic
            Next ir
        
        'Cube 3
            For ir = 1 To 8
                For ic = 1 To 8
                    paDown(3, 3)(ir, ic) = paLeft(3, 1)(ic, (9 - ir))
                Next ic
            Next ir
        
        'Cube 4
            For ir = 1 To 8
                For ic = 1 To 8
                    paDown(3, 4)(ir, ic) = paLeft(3, 3)(ic, (9 - ir))
                Next ic
            Next ir
    
    'Anim: 4 (Regular (same as 2))
    
        For ir = 1 To 8
            For ic = 1 To 8
                'Cube 1
                paDown(4, 1)(ir, ic) = paDown(2, 1)(ir, ic)
                'Cube 2
                paDown(4, 2)(ir, ic) = paDown(2, 2)(ir, ic)
                'Cube 3
                paDown(4, 3)(ir, ic) = paDown(2, 3)(ir, ic)
                'Cube 4
                paDown(4, 4)(ir, ic) = paDown(2, 4)(ir, ic)
            Next ic
        Next ir

End Sub
