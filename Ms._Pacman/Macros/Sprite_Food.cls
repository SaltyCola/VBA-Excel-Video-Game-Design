VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sprite_Food"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'=================================================================================='
'=================================== Properties ==================================='
'=================================================================================='

'AI Module
Private pAIModule As AI_Food 'Holds the AI for the Food objects

'Food Specific Properties
Private pAmbientAudio As Boolean 'Tells the food when to play ambient audio
Private pEaten As Boolean 'Tells the food object to switch to its points anim and then disappear
Private pCounter As Integer 'Tells Food how long to keep points on screen until it disappears
Private pPoints As Integer 'Tells Score how much this food item is worth
Private pAnimExtender As Integer 'Allows to slow down animation

'Coordinates (Top-Left Corner of C1)
Private pXcol As Integer 'x, or column, coordinate of the top-left-most pixel
Private pYrow As Integer 'y, or row, coordinate of the top-left-most pixel

'Velocity
Private pSpeed As Integer 'speed of Food
Private pVx As Integer 'x, or column, velocity of Food
Private pVy As Integer 'y, or row, velocity of Food

'Animation
Private pnDir As Integer '1=Up ; 2=Left ; 3=Right ; 4=Down
Private pnAnim As Integer '1=Bounce up ; 2=Bounce down

'Current Cubes for drawing
Private pc1 As Cube 'top left cube
Private pc2 As Cube 'rop right cube
Private pc3 As Cube 'bottom left cube
Private pc4 As Cube 'bottom right cube

'Sprite Arrays (READ ONLY) (1 to 3 [anims], 1 to 4 [cubes]) each cube holds a 8x8 array
Private paFood(1 To 3, 1 To 4) As Variant 'holds all 4 cubes' pixel colors for each of the 3 frames of the Food animation (3rd frame is the points frame)

'''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
' AI Module
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Property Get AIModule() As AI_Food
    Set AIModule = pAIModule
End Property
Public Property Let AIModule(Value As AI_Food)
    Set pAIModule = Value
End Property

'''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
' Ambient Audio Boolean
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Property Get AmbientAudio() As Boolean
    AmbientAudio = pAmbientAudio
End Property
Public Property Let AmbientAudio(Value As Boolean)
    pAmbientAudio = Value
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
' Counter
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Property Get Counter() As Integer
    Counter = pCounter
End Property
Public Property Let Counter(Value As Integer)
    pCounter = Value
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
' Animation Extender
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Property Get AnimExtender() As Integer
    AnimExtender = pAnimExtender
End Property
Public Property Let AnimExtender(Value As Integer)
    pAnimExtender = Value
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
' Direction Number (1=Up)
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
' Animation Number (1=Bounce down ; 2=Bounce up)
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
' Food Anim Array
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Property Get aFood(iAnimNum As Integer, iCube As Integer) As Variant
    aFood = paFood(iAnimNum, iCube)
End Property
'Public Property Let aFood(iAnimNum As Integer, iCube As Integer, Value As Variant)
'    paFood(iAnimNum, iCube) = Value
'End Property


'================================================================================='
'================================ Public Methods ================================='
'================================================================================='

Public Sub InitializeSprites(ByVal FoodName As String)
'initialize all animation cubes for current "FoodName"

    'intialize Anim Arrays
    If LCase(FoodName) = "blank" Then
        Call SetAnimArray_Blank
    ElseIf LCase(FoodName) = "cherry" Then
        Call SetAnimArray_Cherry
    ElseIf LCase(FoodName) = "strawberry" Then
        Call SetAnimArray_Strawberry
    ElseIf LCase(FoodName) = "peach" Then
        Call SetAnimArray_Peach
    ElseIf LCase(FoodName) = "pretzel" Then
        Call SetAnimArray_Pretzel
    ElseIf LCase(FoodName) = "apple" Then
        Call SetAnimArray_Apple
    ElseIf LCase(FoodName) = "pear" Then
        Call SetAnimArray_Pear
    ElseIf LCase(FoodName) = "banana" Then
        Call SetAnimArray_Banana
    ElseIf LCase(FoodName) = "graduation hat please" Then
        MsgBox "Shut up Emily."
    End If

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

Public Sub UpdatePosition(Optional FoodTypesArrCall As Boolean)

    Dim bU As Boolean 'is Up turn allowed
    Dim bL As Boolean 'is Left turn allowed
    Dim bR As Boolean 'is Right turn allowed
    Dim bD As Boolean 'is Down turn allowed
    Dim sprtRngU As Range 'sprite's potential pos after dir change up
    Dim sprtRngL As Range 'sprite's potential pos after dir change left
    Dim sprtRngR As Range 'sprite's potential pos after dir change right
    Dim sprtRngD As Range 'sprite's potential pos after dir change down
    
    'collision with MsP
    If Not FoodTypesArrCall Then 'if the call did not come from the food types array
        'food close enough to be eaten
        If Not pEaten And (Abs(MsP.Xcol - pXcol) <= 2) And (Abs(MsP.Yrow - pYrow) <= 2) Then
            'change pacdot to eaten
            pEaten = True
            'Update Scores
            ScoreUpdateBool = True
            Call UpdateScores(Me.Points)
            'stop and reset Ambient audio
            Call StopAudio(arrFoodAmbient, True, "A")
            Call StopAudio(arrFoodAmbient, True, "B")
            'change Ambient Audio boolean
            pAmbientAudio = False
            'play Food Eat audio
            Call PlayAudio(arrFoodEat, False)
        End If
    End If
    
'=============================================================================================================================================
    'AI Decision Making and Movement
    If Not pEaten Then
        'initialize directional turn booleans
            bU = False
            bL = False
            bR = False
            bD = False
        'decrement timer
            If pAIModule.Timer <> -1 Then: pAIModule.TimerDecrement
        'timer hits 0
            If pAIModule.Timer = 0 Then
                'choose random entrance and exit tunnels
                pAIModule.TunnelsRandom Me
                'start ambient audio
                pAmbientAudio = True
        'Sprite is on the map
            ElseIf pAIModule.Timer = -1 Then
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
                    pAIModule.DetermineDirection Me, bU, bL, bR, bD
                'Move sprite to new coordinates
                    pXcol = pXcol + (pVx * pSpeed)
                    pYrow = pYrow + (pVy * pSpeed)
            End If
        'target reached
            If pXcol = pAIModule.Xtarget And pYrow = pAIModule.Ytarget Then
                'food has exited without being eaten
                If pAIModule.TargetCounter = 7 Then
                    'hide food
                    Me.HideFood
                'change targets
                Else
                    pAIModule.TargetChange
                End If
            End If
    End If
'=============================================================================================================================================
    
    'update pixel colors
        If pEaten Then
            pnAnim = 3 'Points anim
            'decrement points counter
            pCounter = pCounter - 1
            'points counter hits 0
            If pCounter = 0 Then
                'blank out and move food to waiting position
                    Me.HideFood
                    Me.SetPosition
                    Me.InitializeSprites ("Blank")
                'change audio playing boolean
                    arrFoodEat(2) = False
            End If
        ElseIf Not pEaten And pnAnim = 1 Then
            pAnimExtender = pAnimExtender - 1
            If pAnimExtender = 0 Then
                pAnimExtender = 2
                pnAnim = 2
            End If
        ElseIf Not pEaten And pnAnim = 2 Then
            pAnimExtender = pAnimExtender - 1
            If pAnimExtender = 0 Then
                pAnimExtender = 2
                pnAnim = 1
            End If
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

    Dim ir As Integer 'integer iterator for pixel rows of cube
    Dim ic As Integer 'integer iterator for pixel columns of cube
    
    'set cubes
    For ir = 1 To 8
        For ic = 1 To 8
            pc1.Pixels(ir, ic).Color = paFood(pnAnim, 1)(ir, ic)
            pc2.Pixels(ir, ic).Color = paFood(pnAnim, 2)(ir, ic)
            pc3.Pixels(ir, ic).Color = paFood(pnAnim, 3)(ir, ic)
            pc4.Pixels(ir, ic).Color = paFood(pnAnim, 4)(ir, ic)
        Next ic
    Next ir

End Sub

Public Sub Draw()
'Draws this Ghost at its current coordinates

    'Draw Cubes in order
    pc1.Draw
    pc2.Draw
    pc3.Draw
    pc4.Draw

End Sub

Public Sub HideFood()
'Place food object back into waiting position and stop ambient audio.

    'Stop Movement
    pVx = 0
    pVy = 0
    
    'change x and y coords
    pXcol = 241
    pYrow = 273
    
    'reset food ai timer
    If Not pEaten Then: pAIModule.TimerReset
    
    'stop playing ambient audio
    pAmbientAudio = False

End Sub


'=================================================================================='
'================================ Private Methods ================================='
'=================================================================================='

Private Sub Class_Initialize()

    Dim arrTemp(1 To 8, 1 To 8) As Long 'temporary array to save space for pixel colors
    Dim i1 As Integer 'integer iterator
    Dim i2 As Integer 'integer iterator
    
    'initialize AI
    Set pAIModule = New AI_Food
    pAIModule.TimerReset
    
    'initialize properties
    pAmbientAudio = False 'starts with no audio playing
    pEaten = False 'starts not eaten
    pCounter = 20 'This means the points anim will remain on screen for 20 frames
    pPoints = 0 'points value is decided in each anim initializer
    pAnimExtender = 2
    pnDir = 1 'starts facing up
    pnAnim = 1 'starts with animation frame 1: Bounce Up
    pSpeed = 2 'starts with 4 pixel speed
    pVx = 0 'starts with zero velocity
    pVy = 0 'starts with zero velocity
    
    'initialize cubes
    Set pc1 = New Cube
    Set pc2 = New Cube
    Set pc3 = New Cube
    Set pc4 = New Cube
    
    'initialize arrays
    For i1 = 1 To 3
        For i2 = 1 To 4
            'Food anim array
            paFood(i1, i2) = arrTemp
        Next i2
    Next i1
    
    'initialize original coordinates
    Me.HideFood 'waiting position
    Me.SetPosition 'set food position

End Sub

Private Sub SetAnimArray_Blank()

    Dim ia As Integer 'anim iterator
    Dim ib As Integer 'cube iterator
    Dim ir As Integer 'generic row iteration integer
    Dim ic As Integer 'generic column iteration integer
    
    'set food points
    pPoints = 0
    
    For ia = 1 To 3
        For ib = 1 To 4
            For ir = 1 To 8
                For ic = 1 To 8
                    paFood(ia, ib)(ir, ic) = Black
                Next ic
            Next ir
        Next ib
    Next ia

End Sub

Private Sub SetAnimArray_Cherry()

    Dim ir As Integer 'generic row iteration integer
    Dim ic As Integer 'generic column iteration integer
    
    'set food points
    pPoints = 100
    
    'Anim: 1 (Bounce Up)
        'Cube 1
            'row 5
                paFood(1, 1)(5, 8) = Brown
            'row 6
            For ic = 4 To 6
                paFood(1, 1)(6, ic) = Red
            Next ic
                paFood(1, 1)(6, 7) = Brown
            'row 7
            For ic = 3 To 8
                paFood(1, 1)(7, ic) = Red
            Next ic
                paFood(1, 1)(7, 6) = Brown
            'row 8
            For ic = 3 To 7
                paFood(1, 1)(8, ic) = Red
            Next ic
        'Cube 2
            'row 2
                paFood(1, 2)(2, 5) = Brown
                paFood(1, 2)(2, 6) = Brown
            'row 3
            For ic = 3 To 6
                paFood(1, 2)(3, ic) = Brown
            Next ic
            'row 4
                paFood(1, 2)(4, 1) = Brown
                paFood(1, 2)(4, 2) = Brown
                paFood(1, 2)(4, 4) = Brown
            'row 5
                paFood(1, 2)(5, 4) = Brown
            'row 6
                paFood(1, 2)(6, 3) = Brown
            'row 7
                paFood(1, 2)(7, 2) = Brown
            'row 8
            For ic = 1 To 4
                paFood(1, 2)(8, ic) = Red
            Next ic
                paFood(1, 2)(8, 2) = Brown
        'Cube 3
            'base red rows 1-4
            For ir = 1 To 3
                For ic = 3 To 8
                    paFood(1, 3)(ir, ic) = Red
                Next ic
            Next ir
                    paFood(1, 3)(4, 8) = Red
            'black and white
                'row 1
                    paFood(1, 3)(1, 4) = White
                    paFood(1, 3)(1, 7) = Black
                'row 2
                    paFood(1, 3)(2, 5) = White
                    paFood(1, 3)(2, 7) = Black
                'row 3
                    paFood(1, 3)(3, 3) = Black
                    paFood(1, 3)(3, 7) = Black
        'Cube 4
            'base red rows 1-5
            For ir = 1 To 5
                For ic = 1 To 5
                    paFood(1, 4)(ir, ic) = Red
                Next ic
            Next ir
            'white black and browns
                'row 1
                    paFood(1, 4)(1, 2) = Brown
                'row 3
                    paFood(1, 4)(3, 1) = White
                'row 4
                    paFood(1, 4)(4, 2) = White
                'row 5
                    paFood(1, 4)(5, 5) = Black
    
    'Anim: 2 (Bounce Down) (anim2 = anim1 except shifted down 2 pixels)
        'Cube 1
            'rows 3-8
            For ir = 3 To 8
                For ic = 1 To 8
                    paFood(2, 1)(ir, ic) = paFood(1, 1)((ir - 2), ic)
                Next ic
            Next ir
        'Cube 2
            'rows 3-8
            For ir = 3 To 8
                For ic = 1 To 8
                    paFood(2, 2)(ir, ic) = paFood(1, 2)((ir - 2), ic)
                Next ic
            Next ir
        'Cube 3
            'rows 1-2 (last two rows from anim 1 cube 1)
            For ir = 1 To 2
                For ic = 1 To 8
                    paFood(2, 3)(ir, ic) = paFood(1, 1)((ir + 6), ic)
                Next ic
            Next ir
            'rows 3-8
            For ir = 3 To 8
                For ic = 1 To 8
                    paFood(2, 3)(ir, ic) = paFood(1, 3)((ir - 2), ic)
                Next ic
            Next ir
        'Cube 4
            'rows 1-2 (last two rows from anim 1 cube 2)
            For ir = 1 To 2
                For ic = 1 To 8
                    paFood(2, 4)(ir, ic) = paFood(1, 2)((ir + 6), ic)
                Next ic
            Next ir
            'rows 3-8
            For ir = 3 To 8
                For ic = 1 To 8
                    paFood(2, 4)(ir, ic) = paFood(1, 4)((ir - 2), ic)
                Next ic
            Next ir
    
    'Anim: 3 (Points)
        'Cube 1
            'rows 3-7
            For ir = 3 To 7
                paFood(3, 1)(ir, 4) = White
            Next ir
            'row 6
                paFood(3, 1)(6, 8) = White
            'rows 6-8
            For ir = 6 To 8
                paFood(3, 1)(ir, 7) = White
            Next ir
        'Cube 2
            'rows 6-8
            For ir = 6 To 8
                paFood(3, 2)(ir, 1) = White
            Next ir
        'Cube 3
            'row 1
                paFood(3, 3)(1, 7) = White
            'row 2
            For ic = 7 To 8
                paFood(3, 3)(2, ic) = White
            Next ic
        'Cube 4
            'row 1
                paFood(3, 4)(1, 4) = White
            'rows 1-2
            For ir = 1 To 2
                paFood(3, 4)(ir, 1) = White
            Next ir
            'rows 1-5
            For ir = 1 To 5
                paFood(3, 4)(ir, 3) = White
                paFood(3, 4)(ir, 5) = White
            Next ir
            'row 5
                paFood(3, 4)(5, 4) = White

End Sub

Private Sub SetAnimArray_Strawberry()

    Dim ir As Integer 'generic row iteration integer
    Dim ic As Integer 'generic column iteration integer
    
    'set food points
    pPoints = 200
    
    'Anim: 1 (Bounce Up)
        'Cube 1
            'Red pixels
                'rows 4-8
                For ir = 4 To 8
                    For ic = 3 To 8
                        paFood(1, 1)(ir, ic) = Red
                    Next ic
                Next ir
            'Green pixels
                'row 3
                For ic = 5 To 7
                    paFood(1, 1)(3, ic) = Green
                Next ic
                'row 4
                For ic = 6 To 8
                    paFood(1, 1)(4, ic) = Green
                Next ic
                'row 5
                    paFood(1, 1)(5, 8) = Green
            'White pixels
                    paFood(1, 1)(2, 8) = White
                    paFood(1, 1)(3, 8) = White
                    paFood(1, 1)(6, 4) = White
                    paFood(1, 1)(7, 6) = White
                    paFood(1, 1)(7, 8) = White
            'Black pixels
                    paFood(1, 1)(4, 3) = Black
        'Cube 2
            'Red pixels
                'rows 4-8
                For ir = 4 To 8
                    For ic = 1 To 5
                        paFood(1, 2)(ir, ic) = Red
                    Next ic
                Next ir
            'Green pixels
                'row 3
                For ic = 1 To 3
                    paFood(1, 2)(3, ic) = Green
                Next ic
                'row 4
                For ic = 1 To 2
                    paFood(1, 2)(4, ic) = Green
                Next ic
            'White pixels
                    paFood(1, 2)(5, 4) = White
                    paFood(1, 2)(6, 2) = White
                    paFood(1, 2)(8, 3) = White
            'Black pixels
                    paFood(1, 2)(4, 5) = Black
        'Cube 3
            'Red pixels
                'rows 1-2
                For ir = 1 To 2
                    For ic = 4 To 8
                        paFood(1, 3)(ir, ic) = Red
                    Next ic
                Next ir
                'row 3
                For ic = 5 To 8
                    paFood(1, 3)(3, ic) = Red
                Next ic
                'row 4
                For ic = 6 To 8
                    paFood(1, 3)(4, ic) = Red
                Next ic
                'row 5
                    paFood(1, 3)(5, 8) = Red
            'White pixels
                paFood(1, 3)(1, 5) = White
                paFood(1, 3)(1, 8) = White
                paFood(1, 3)(3, 7) = White
        'Cube 4
            'Red pixels
                'rows 1-2
                For ir = 1 To 2
                    For ic = 1 To 4
                        paFood(1, 4)(ir, ic) = Red
                    Next ic
                Next ir
                'row 3
                For ic = 1 To 3
                    paFood(1, 4)(3, ic) = Red
                Next ic
                'row 4
                For ic = 1 To 2
                    paFood(1, 4)(4, ic) = Red
                Next ic
            'White pixels
                paFood(1, 4)(3, 2) = White
    
    'Anim: 2 (Bounce Down) (anim2 = anim1 except shifted down 2 pixels)
        'Cube 1
            'rows 3-8
            For ir = 3 To 8
                For ic = 1 To 8
                    paFood(2, 1)(ir, ic) = paFood(1, 1)((ir - 2), ic)
                Next ic
            Next ir
        'Cube 2
            'rows 3-8
            For ir = 3 To 8
                For ic = 1 To 8
                    paFood(2, 2)(ir, ic) = paFood(1, 2)((ir - 2), ic)
                Next ic
            Next ir
        'Cube 3
            'rows 1-2 (last two rows from anim 1 cube 1)
            For ir = 1 To 2
                For ic = 1 To 8
                    paFood(2, 3)(ir, ic) = paFood(1, 1)((ir + 6), ic)
                Next ic
            Next ir
            'rows 3-8
            For ir = 3 To 8
                For ic = 1 To 8
                    paFood(2, 3)(ir, ic) = paFood(1, 3)((ir - 2), ic)
                Next ic
            Next ir
        'Cube 4
            'rows 1-2 (last two rows from anim 1 cube 2)
            For ir = 1 To 2
                For ic = 1 To 8
                    paFood(2, 4)(ir, ic) = paFood(1, 2)((ir + 6), ic)
                Next ic
            Next ir
            'rows 3-8
            For ir = 3 To 8
                For ic = 1 To 8
                    paFood(2, 4)(ir, ic) = paFood(1, 4)((ir - 2), ic)
                Next ic
            Next ir
    
    'Anim: 3 (Points)
        'Cube 1
            'rows 3-7
            For ir = 3 To 7
                For ic = 3 To 5
                    paFood(3, 1)(ir, ic) = White
                Next ic
            Next ir
                    paFood(3, 1)(4, 3) = Black
                    paFood(3, 1)(4, 4) = Black
                    paFood(3, 1)(6, 4) = Black
                    paFood(3, 1)(6, 5) = Black
            'row 6
                paFood(3, 1)(6, 8) = White
            'rows 6-8
            For ir = 6 To 8
                paFood(3, 1)(ir, 7) = White
            Next ir
        'Cube 2
            'rows 6-8
            For ir = 6 To 8
                paFood(3, 2)(ir, 1) = White
            Next ir
        'Cube 3
            'row 1
                paFood(3, 3)(1, 7) = White
            'row 2
            For ic = 7 To 8
                paFood(3, 3)(2, ic) = White
            Next ic
        'Cube 4
            'row 1
                paFood(3, 4)(1, 4) = White
            'rows 1-2
            For ir = 1 To 2
                paFood(3, 4)(ir, 1) = White
            Next ir
            'rows 1-5
            For ir = 1 To 5
                paFood(3, 4)(ir, 3) = White
                paFood(3, 4)(ir, 5) = White
            Next ir
            'row 5
                paFood(3, 4)(5, 4) = White

End Sub

Private Sub SetAnimArray_Peach()

    Dim ir As Integer 'generic row iteration integer
    Dim ic As Integer 'generic column iteration integer
    
    'set food points
    pPoints = 500
    
    'Anim: 1 (Bounce Up)
        'Cube 1
            'Orange pixels
                'row 5
                For ic = 5 To 8
                    paFood(1, 1)(5, ic) = Orange
                Next ic
                'row 6
                For ic = 4 To 8
                    paFood(1, 1)(6, ic) = Orange
                Next ic
                'row 7-8
                For ir = 7 To 8
                    For ic = 3 To 8
                        paFood(1, 1)(ir, ic) = Orange
                    Next ic
                Next ir
        'Cube 2
            'y-axis reflection of Cube 1
                'rows 5-8
                For ir = 5 To 8
                    For ic = 1 To 8
                        paFood(1, 2)(ir, ic) = paFood(1, 1)(ir, (9 - ic))
                    Next ic
                Next ir
            'Green pixels
                'row 2
                    paFood(1, 2)(2, 2) = Green
                    paFood(1, 2)(2, 3) = Green
                'row 3
                For ic = 1 To 5
                    paFood(1, 2)(3, ic) = Green
                Next ic
                'row 4
                For ic = 2 To 4
                    paFood(1, 2)(4, ic) = Green
                Next ic
            'Brown pixels
                paFood(1, 2)(5, 1) = Brown
        'Cube 1(stem)
            'rows 3-6
            For ir = 3 To 6
                paFood(1, 1)(ir, 8) = Brown
            Next ir
            'row 5
                paFood(1, 1)(5, 7) = Brown
        'Cube 3
            'Orange pixels
                'rows 1-5
                For ir = 1 To 5
                    For ic = 3 To 8
                        paFood(1, 3)(ir, ic) = Orange
                    Next ic
                Next ir
            'Black pixels
                'rows 3-4
                For ir = 3 To 4
                    paFood(1, 3)(ir, 3) = Black
                Next ir
                'row 5
                For ic = 3 To 5
                    paFood(1, 3)(5, ic) = Black
                Next ic
        'Cube 4
            'y-axis reflection of Cube 3
                'rows 1-5
                For ir = 1 To 5
                    For ic = 1 To 8
                        paFood(1, 4)(ir, ic) = paFood(1, 3)(ir, (9 - ic))
                    Next ic
                Next ir
    
    'Anim: 2 (Bounce Down) (anim2 = anim1 except shifted down 2 pixels)
        'Cube 1
            'rows 3-8
            For ir = 3 To 8
                For ic = 1 To 8
                    paFood(2, 1)(ir, ic) = paFood(1, 1)((ir - 2), ic)
                Next ic
            Next ir
        'Cube 2
            'rows 3-8
            For ir = 3 To 8
                For ic = 1 To 8
                    paFood(2, 2)(ir, ic) = paFood(1, 2)((ir - 2), ic)
                Next ic
            Next ir
        'Cube 3
            'rows 1-2 (last two rows from anim 1 cube 1)
            For ir = 1 To 2
                For ic = 1 To 8
                    paFood(2, 3)(ir, ic) = paFood(1, 1)((ir + 6), ic)
                Next ic
            Next ir
            'rows 3-8
            For ir = 3 To 8
                For ic = 1 To 8
                    paFood(2, 3)(ir, ic) = paFood(1, 3)((ir - 2), ic)
                Next ic
            Next ir
        'Cube 4
            'rows 1-2 (last two rows from anim 1 cube 2)
            For ir = 1 To 2
                For ic = 1 To 8
                    paFood(2, 4)(ir, ic) = paFood(1, 2)((ir + 6), ic)
                Next ic
            Next ir
            'rows 3-8
            For ir = 3 To 8
                For ic = 1 To 8
                    paFood(2, 4)(ir, ic) = paFood(1, 4)((ir - 2), ic)
                Next ic
            Next ir
    
    'Anim: 3 (Points)
        'Cube 1
            'rows 3-7
            For ir = 3 To 7
                For ic = 3 To 5
                    paFood(3, 1)(ir, ic) = White
                Next ic
            Next ir
                    paFood(3, 1)(4, 4) = Black
                    paFood(3, 1)(4, 5) = Black
                    paFood(3, 1)(6, 3) = Black
                    paFood(3, 1)(6, 4) = Black
            'row 6
                paFood(3, 1)(6, 8) = White
            'rows 6-8
            For ir = 6 To 8
                paFood(3, 1)(ir, 7) = White
            Next ir
        'Cube 2
            'rows 6-8
            For ir = 6 To 8
                paFood(3, 2)(ir, 1) = White
            Next ir
        'Cube 3
            'row 1
                paFood(3, 3)(1, 7) = White
            'row 2
            For ic = 7 To 8
                paFood(3, 3)(2, ic) = White
            Next ic
        'Cube 4
            'row 1
                paFood(3, 4)(1, 4) = White
            'rows 1-2
            For ir = 1 To 2
                paFood(3, 4)(ir, 1) = White
            Next ir
            'rows 1-5
            For ir = 1 To 5
                paFood(3, 4)(ir, 3) = White
                paFood(3, 4)(ir, 5) = White
            Next ir
            'row 5
                paFood(3, 4)(5, 4) = White

End Sub

Private Sub SetAnimArray_Pretzel()

    Dim ir As Integer 'generic row iteration integer
    Dim ic As Integer 'generic column iteration integer
    
    'set food points
    pPoints = 700
    
    'Anim: 1 (Bounce Up)
        'Cube 1
            'row 2
                paFood(1, 1)(2, 4) = White
                paFood(1, 1)(2, 5) = Orange
                paFood(1, 1)(2, 6) = Orange
            'row 3
            For ic = 3 To 8
                paFood(1, 1)(3, ic) = Orange
            Next ic
            'row 4
                paFood(1, 1)(4, 3) = Orange
                paFood(1, 1)(4, 4) = Orange
                paFood(1, 1)(4, 7) = Orange
            'row 5
                paFood(1, 1)(5, 2) = Orange
                paFood(1, 1)(5, 3) = Orange
            'row 6
                paFood(1, 1)(6, 2) = Orange
                paFood(1, 1)(6, 3) = Orange
                paFood(1, 1)(6, 8) = Orange
            'row 7
                paFood(1, 1)(7, 2) = Orange
                paFood(1, 1)(7, 3) = Orange
                paFood(1, 1)(7, 8) = Orange
            'row 8
                paFood(1, 1)(8, 2) = Orange
                paFood(1, 1)(8, 3) = Orange
                paFood(1, 1)(8, 4) = White
                paFood(1, 1)(8, 7) = White
                paFood(1, 1)(8, 8) = Orange
        'Cube 2
            'row 2
                paFood(1, 2)(2, 2) = White
                paFood(1, 2)(2, 4) = Orange
                paFood(1, 2)(2, 5) = Orange
            'row 3
            For ic = 2 To 6
                paFood(1, 2)(3, ic) = Orange
            Next ic
            'row 4
                paFood(1, 2)(4, 2) = Orange
                paFood(1, 2)(4, 3) = Orange
                paFood(1, 2)(4, 5) = White
                paFood(1, 2)(4, 6) = Orange
                paFood(1, 2)(4, 7) = Orange
            'row 5
                paFood(1, 2)(5, 1) = Orange
                paFood(1, 2)(5, 2) = Orange
                paFood(1, 2)(5, 6) = Orange
                paFood(1, 2)(5, 7) = Orange
            'row 6
                paFood(1, 2)(6, 1) = Orange
                paFood(1, 2)(6, 6) = Orange
                paFood(1, 2)(6, 7) = Orange
            'row 7
                paFood(1, 2)(7, 1) = Orange
                paFood(1, 2)(7, 6) = Orange
                paFood(1, 2)(7, 7) = Orange
            'row 8
                paFood(1, 2)(8, 2) = Orange
                paFood(1, 2)(8, 6) = Orange
                paFood(1, 2)(8, 7) = Orange
        'Cube 3
            'row 1
                paFood(1, 3)(1, 3) = Orange
                paFood(1, 3)(1, 4) = Orange
                paFood(1, 3)(1, 7) = Orange
                paFood(1, 3)(1, 8) = Orange
            'row 2
                paFood(1, 3)(2, 4) = Orange
                paFood(1, 3)(2, 6) = Orange
                paFood(1, 3)(2, 7) = Orange
            'row 3
                paFood(1, 3)(3, 6) = Orange
                paFood(1, 3)(3, 7) = Orange
            'row 4
                paFood(1, 3)(4, 5) = Orange
                paFood(1, 3)(4, 6) = Orange
                paFood(1, 3)(4, 8) = Orange
            'row 5
                paFood(1, 3)(5, 4) = Orange
                paFood(1, 3)(5, 5) = Orange
                paFood(1, 3)(5, 8) = Orange
            'row 6
                paFood(1, 3)(6, 3) = Orange
                paFood(1, 3)(6, 4) = Orange
                paFood(1, 3)(6, 5) = White
        'Cube 4
            'row 1
                paFood(1, 4)(1, 2) = Orange
                paFood(1, 4)(1, 3) = Orange
                paFood(1, 4)(1, 5) = Orange
                paFood(1, 4)(1, 6) = Orange
                paFood(1, 4)(1, 7) = White
            'row 2
                paFood(1, 4)(2, 3) = Orange
                paFood(1, 4)(2, 5) = Orange
                paFood(1, 4)(2, 6) = Orange
            'row 3
                paFood(1, 4)(3, 3) = Orange
                paFood(1, 4)(3, 4) = Orange
            'row 4
                paFood(1, 4)(4, 1) = Orange
                paFood(1, 4)(4, 2) = Orange
                paFood(1, 4)(4, 4) = Orange
                paFood(1, 4)(4, 5) = Orange
            'row 5
                paFood(1, 4)(5, 1) = Orange
                paFood(1, 4)(5, 2) = Orange
                paFood(1, 4)(5, 4) = Orange
                paFood(1, 4)(5, 5) = Orange
                paFood(1, 4)(5, 6) = White
            'row 6
                paFood(1, 4)(6, 5) = Orange
                paFood(1, 4)(6, 6) = Orange
    
    'Anim: 2 (Bounce Down) (anim2 = anim1 except shifted down 2 pixels)
        'Cube 1
            'rows 3-8
            For ir = 3 To 8
                For ic = 1 To 8
                    paFood(2, 1)(ir, ic) = paFood(1, 1)((ir - 2), ic)
                Next ic
            Next ir
        'Cube 2
            'rows 3-8
            For ir = 3 To 8
                For ic = 1 To 8
                    paFood(2, 2)(ir, ic) = paFood(1, 2)((ir - 2), ic)
                Next ic
            Next ir
        'Cube 3
            'rows 1-2 (last two rows from anim 1 cube 1)
            For ir = 1 To 2
                For ic = 1 To 8
                    paFood(2, 3)(ir, ic) = paFood(1, 1)((ir + 6), ic)
                Next ic
            Next ir
            'rows 3-8
            For ir = 3 To 8
                For ic = 1 To 8
                    paFood(2, 3)(ir, ic) = paFood(1, 3)((ir - 2), ic)
                Next ic
            Next ir
        'Cube 4
            'rows 1-2 (last two rows from anim 1 cube 2)
            For ir = 1 To 2
                For ic = 1 To 8
                    paFood(2, 4)(ir, ic) = paFood(1, 2)((ir + 6), ic)
                Next ic
            Next ir
            'rows 3-8
            For ir = 3 To 8
                For ic = 1 To 8
                    paFood(2, 4)(ir, ic) = paFood(1, 4)((ir - 2), ic)
                Next ic
            Next ir
    
    'Anim: 3 (Points)
        'Cube 1
            'rows 3-7
            For ic = 3 To 5
                paFood(3, 1)(3, ic) = White
            Next ic
            For ir = 3 To 7
                paFood(3, 1)(ir, 5) = White
            Next ir
            'row 6
                paFood(3, 1)(6, 8) = White
            'rows 6-8
            For ir = 6 To 8
                paFood(3, 1)(ir, 7) = White
            Next ir
        'Cube 2
            'rows 6-8
            For ir = 6 To 8
                paFood(3, 2)(ir, 1) = White
            Next ir
        'Cube 3
            'row 1
                paFood(3, 3)(1, 7) = White
            'row 2
            For ic = 7 To 8
                paFood(3, 3)(2, ic) = White
            Next ic
        'Cube 4
            'row 1
                paFood(3, 4)(1, 4) = White
            'rows 1-2
            For ir = 1 To 2
                paFood(3, 4)(ir, 1) = White
            Next ir
            'rows 1-5
            For ir = 1 To 5
                paFood(3, 4)(ir, 3) = White
                paFood(3, 4)(ir, 5) = White
            Next ir
            'row 5
                paFood(3, 4)(5, 4) = White

End Sub

Private Sub SetAnimArray_Apple()

    Dim ir As Integer 'generic row iteration integer
    Dim ic As Integer 'generic column iteration integer
    
    'set food points
    pPoints = 1000
    
    'Anim: 1 (Bounce Up)
        'Cube 1
            'Red pixels
                'rows 3-8
                For ir = 3 To 8
                    For ic = 3 To 8
                        paFood(1, 1)(ir, ic) = Red
                    Next ic
                Next ir
            'Brown pixels
                'row 3-4
                    paFood(1, 1)(3, 8) = Brown
                    paFood(1, 1)(4, 8) = Brown
            'Black pixels
                'row 3
                    paFood(1, 1)(3, 3) = Black
                    paFood(1, 1)(3, 7) = Black
        'Cube 2
            'Red pixels
                'rows 3-8
                For ir = 3 To 8
                    For ic = 1 To 6
                        paFood(1, 2)(ir, ic) = Red
                    Next ic
                Next ir
            'Brown pixels
                'row 2
                    paFood(1, 2)(2, 1) = Brown
            'White pixels
                'row 8
                    paFood(1, 2)(8, 4) = White
            'Black pixels
                'row 3-4
                    paFood(1, 2)(3, 1) = Black
                    paFood(1, 2)(3, 5) = Black
                    paFood(1, 2)(3, 6) = Black
                    paFood(1, 2)(4, 6) = Black
        'Cube 3
            'Red pixels
                'rows 1-5
                For ir = 1 To 5
                    For ic = 3 To 8
                        paFood(1, 3)(ir, ic) = Red
                    Next ic
                Next ir
            'Black pixels
                'row 2-5
                    For ir = 2 To 5
                        paFood(1, 3)(ir, 3) = Black
                    Next ir
                    paFood(1, 3)(4, 4) = Black
                    paFood(1, 3)(5, 4) = Black
                    paFood(1, 3)(5, 5) = Black
                    paFood(1, 3)(5, 8) = Black
        'Cube 4
            'Red pixels
                'rows 1-5
                For ir = 1 To 5
                    For ic = 1 To 6
                        paFood(1, 4)(ir, ic) = Red
                    Next ic
                Next ir
            'White pixels
                'rows 1-2
                    paFood(1, 4)(1, 4) = White
                    paFood(1, 4)(2, 3) = White
            'Black pixels
                'row 2-5
                    For ir = 2 To 5
                        paFood(1, 4)(ir, 6) = Black
                    Next ir
                    paFood(1, 4)(4, 5) = Black
                    paFood(1, 4)(5, 5) = Black
                    paFood(1, 4)(5, 4) = Black
    
    'Anim: 2 (Bounce Down) (anim2 = anim1 except shifted down 2 pixels)
        'Cube 1
            'rows 3-8
            For ir = 3 To 8
                For ic = 1 To 8
                    paFood(2, 1)(ir, ic) = paFood(1, 1)((ir - 2), ic)
                Next ic
            Next ir
        'Cube 2
            'rows 3-8
            For ir = 3 To 8
                For ic = 1 To 8
                    paFood(2, 2)(ir, ic) = paFood(1, 2)((ir - 2), ic)
                Next ic
            Next ir
        'Cube 3
            'rows 1-2 (last two rows from anim 1 cube 1)
            For ir = 1 To 2
                For ic = 1 To 8
                    paFood(2, 3)(ir, ic) = paFood(1, 1)((ir + 6), ic)
                Next ic
            Next ir
            'rows 3-8
            For ir = 3 To 8
                For ic = 1 To 8
                    paFood(2, 3)(ir, ic) = paFood(1, 3)((ir - 2), ic)
                Next ic
            Next ir
        'Cube 4
            'rows 1-2 (last two rows from anim 1 cube 2)
            For ir = 1 To 2
                For ic = 1 To 8
                    paFood(2, 4)(ir, ic) = paFood(1, 2)((ir + 6), ic)
                Next ic
            Next ir
            'rows 3-8
            For ir = 3 To 8
                For ic = 1 To 8
                    paFood(2, 4)(ir, ic) = paFood(1, 4)((ir - 2), ic)
                Next ic
            Next ir
    
    'Anim: 3 (Points)
        'Cube 1
            'rows 3-7
            For ir = 3 To 7
                paFood(3, 1)(ir, 4) = White
            Next ir
            'row 5
                paFood(3, 1)(5, 7) = White
            'rows 5-8
            For ir = 5 To 8
                paFood(3, 1)(ir, 6) = White
                paFood(3, 1)(ir, 8) = White
            Next ir
        'Cube 2
            'row 7
                paFood(3, 2)(7, 2) = White
            'rows 7-8
            For ir = 7 To 8
                paFood(3, 2)(ir, 1) = White
                paFood(3, 2)(ir, 3) = White
            Next ir
        'Cube 3
            'row 1
            For ic = 6 To 8
                paFood(3, 3)(1, ic) = White
            Next ic
        'Cube 4
            'row 1,3,5
                paFood(3, 4)(1, 5) = White
                paFood(3, 4)(3, 2) = White
                paFood(3, 4)(5, 5) = White
            'rows 1-3
            For ir = 1 To 3
                paFood(3, 4)(ir, 1) = White
                paFood(3, 4)(ir, 3) = White
            Next ir
            'rows 1-5
            For ir = 1 To 5
                paFood(3, 4)(ir, 4) = White
                paFood(3, 4)(ir, 6) = White
            Next ir

End Sub

Private Sub SetAnimArray_Pear()

    Dim ir As Integer 'generic row iteration integer
    Dim ic As Integer 'generic column iteration integer
    
    'set food points
    pPoints = 2000
    
    'Anim: 1 (Bounce Up)
        'Cube 1
            'rows 3,4,8
                paFood(1, 1)(3, 8) = Green
                paFood(1, 1)(4, 8) = Green
                paFood(1, 1)(8, 6) = Green
            'rows 5-8
            For ir = 5 To 8
                paFood(1, 1)(ir, 7) = Green
                paFood(1, 1)(ir, 8) = Green
            Next ir
        'Cube 2
            'Green pixels
                'rows 3-8
                For ir = 3 To 8
                    For ic = 1 To 3
                        paFood(1, 2)(ir, ic) = Green
                    Next ic
                Next ir
                'row 8
                    paFood(1, 2)(8, 4) = Green
            'White pixels
                'rows 1-2
                    paFood(1, 2)(1, 2) = White
                    paFood(1, 2)(1, 3) = White
                    paFood(1, 2)(2, 1) = White
            'Black pixels
                'row 3-4
                    paFood(1, 2)(3, 3) = Black
                    paFood(1, 2)(4, 3) = Black
        'Cube 3
            'Green pixels
                'rows 1-6
                For ir = 1 To 6
                    For ic = 5 To 8
                        paFood(1, 3)(ir, ic) = Green
                    Next ic
                Next ir
            'Black pixels
                'row 1,5,6
                    paFood(1, 3)(1, 5) = Black
                    paFood(1, 3)(5, 5) = Black
                    paFood(1, 3)(6, 5) = Black
                    paFood(1, 3)(6, 6) = Black
            'Blue pixels
                'row 2,3,4
                    paFood(1, 3)(2, 7) = Blue
                    paFood(1, 3)(3, 7) = Blue
                    paFood(1, 3)(4, 8) = Blue
        'Cube 4
            'Green pixels
                'rows 1-6
                For ir = 1 To 6
                    For ic = 1 To 5
                        paFood(1, 4)(ir, ic) = Green
                    Next ic
                Next ir
            'Black pixels
                'row 1,5,6
                    paFood(1, 4)(1, 5) = Black
                    paFood(1, 4)(5, 5) = Black
                    paFood(1, 4)(6, 5) = Black
                    paFood(1, 4)(6, 4) = Black
    
    'Anim: 2 (Bounce Down) (anim2 = anim1 except shifted down 2 pixels)
        'Cube 1
            'rows 3-8
            For ir = 3 To 8
                For ic = 1 To 8
                    paFood(2, 1)(ir, ic) = paFood(1, 1)((ir - 2), ic)
                Next ic
            Next ir
        'Cube 2
            'rows 3-8
            For ir = 3 To 8
                For ic = 1 To 8
                    paFood(2, 2)(ir, ic) = paFood(1, 2)((ir - 2), ic)
                Next ic
            Next ir
        'Cube 3
            'rows 1-2 (last two rows from anim 1 cube 1)
            For ir = 1 To 2
                For ic = 1 To 8
                    paFood(2, 3)(ir, ic) = paFood(1, 1)((ir + 6), ic)
                Next ic
            Next ir
            'rows 3-8
            For ir = 3 To 8
                For ic = 1 To 8
                    paFood(2, 3)(ir, ic) = paFood(1, 3)((ir - 2), ic)
                Next ic
            Next ir
        'Cube 4
            'rows 1-2 (last two rows from anim 1 cube 2)
            For ir = 1 To 2
                For ic = 1 To 8
                    paFood(2, 4)(ir, ic) = paFood(1, 2)((ir + 6), ic)
                Next ic
            Next ir
            'rows 3-8
            For ir = 3 To 8
                For ic = 1 To 8
                    paFood(2, 4)(ir, ic) = paFood(1, 4)((ir - 2), ic)
                Next ic
            Next ir
    
    'Anim: 3 (Points)
        'Cube 1
            'White pixels
                'rows 3-7
                For ir = 3 To 7
                    For ic = 3 To 5
                        paFood(3, 1)(ir, ic) = White
                    Next ic
                Next ir
                'row 5
                    paFood(3, 1)(5, 7) = White
                'rows 5-8
                For ir = 5 To 8
                    paFood(3, 1)(ir, 6) = White
                    paFood(3, 1)(ir, 8) = White
                Next ir
            'Black pixels
                'row 4,6
                    paFood(3, 1)(4, 3) = Black
                    paFood(3, 1)(4, 4) = Black
                    paFood(3, 1)(6, 4) = Black
                    paFood(3, 1)(6, 5) = Black
        'Cube 2
            'row 7
                paFood(3, 2)(7, 2) = White
            'rows 7-8
            For ir = 7 To 8
                paFood(3, 2)(ir, 1) = White
                paFood(3, 2)(ir, 3) = White
            Next ir
        'Cube 3
            'row 1
            For ic = 6 To 8
                paFood(3, 3)(1, ic) = White
            Next ic
        'Cube 4
            'row 1,3,5
                paFood(3, 4)(1, 5) = White
                paFood(3, 4)(3, 2) = White
                paFood(3, 4)(5, 5) = White
            'rows 1-3
            For ir = 1 To 3
                paFood(3, 4)(ir, 1) = White
                paFood(3, 4)(ir, 3) = White
            Next ir
            'rows 1-5
            For ir = 1 To 5
                paFood(3, 4)(ir, 4) = White
                paFood(3, 4)(ir, 6) = White
            Next ir

End Sub

Private Sub SetAnimArray_Banana()

    Dim ir As Integer 'generic row iteration integer
    Dim ic As Integer 'generic column iteration integer
    
    'set food points
    pPoints = 5000
    
    'Anim: 1 (Bounce Up)
        'Cube 1
            'Yellow pixels
                'row 8
                    paFood(1, 1)(8, 8) = Yellow
        'Cube 2
            'Yellow pixels
                'row 1
                    paFood(1, 2)(1, 5) = Yellow
                'row 2
                    paFood(1, 2)(2, 4) = Yellow
                'row 3
                    paFood(1, 2)(3, 5) = Yellow
                    paFood(1, 2)(3, 6) = Yellow
                'row 4-5
                For ir = 4 To 5
                    For ic = 3 To 6
                        paFood(1, 2)(ir, ic) = Yellow
                    Next ic
                Next ir
                'row 6
                    paFood(1, 2)(6, 2) = Yellow
                    paFood(1, 2)(6, 4) = Yellow
                    paFood(1, 2)(6, 5) = Yellow
                    paFood(1, 2)(6, 6) = Yellow
                'row 7
                    paFood(1, 2)(7, 1) = Yellow
                    paFood(1, 2)(7, 3) = Yellow
                    paFood(1, 2)(7, 4) = Yellow
                    paFood(1, 2)(7, 5) = Yellow
                'row 8
                    paFood(1, 2)(8, 2) = Yellow
                    paFood(1, 2)(8, 3) = Yellow
                    paFood(1, 2)(8, 4) = Yellow
                    paFood(1, 2)(8, 5) = Yellow
        'Cube 3
            'Yellow pixels
                'row 1
                    paFood(1, 3)(1, 7) = Yellow
                'row 2
                    paFood(1, 3)(2, 6) = Yellow
                    paFood(1, 3)(2, 8) = Yellow
                'row 3
                    paFood(1, 3)(3, 5) = Yellow
                    paFood(1, 3)(3, 7) = Yellow
                    paFood(1, 3)(3, 8) = Yellow
                'row 4
                    paFood(1, 3)(4, 3) = Yellow
                'rows 4-5
                For ir = 4 To 5
                    For ic = 4 To 8
                        paFood(1, 3)(ir, ic) = Yellow
                    Next ic
                Next ir
            'White pixels
                'row 4
                    paFood(1, 3)(4, 8) = White
        'Cube 4
            'Yellow pixels
                'rows 1-4
                For ir = 1 To 4
                    For ic = 1 To 4
                        paFood(1, 4)(ir, ic) = Yellow
                    Next ic
                Next ir
            'White pixels
                'row 2-3
                    paFood(1, 4)(2, 2) = White
                    paFood(1, 4)(3, 1) = White
            'Black pixels
                'row 3-4
                    paFood(1, 4)(3, 4) = Black
                    paFood(1, 4)(4, 4) = Black
                    paFood(1, 4)(4, 3) = Black
    
    'Anim: 2 (Bounce Down) (anim2 = anim1 except shifted down 2 pixels)
        'Cube 1
            'rows 3-8
            For ir = 3 To 8
                For ic = 1 To 8
                    paFood(2, 1)(ir, ic) = paFood(1, 1)((ir - 2), ic)
                Next ic
            Next ir
        'Cube 2
            'rows 3-8
            For ir = 3 To 8
                For ic = 1 To 8
                    paFood(2, 2)(ir, ic) = paFood(1, 2)((ir - 2), ic)
                Next ic
            Next ir
        'Cube 3
            'rows 1-2 (last two rows from anim 1 cube 1)
            For ir = 1 To 2
                For ic = 1 To 8
                    paFood(2, 3)(ir, ic) = paFood(1, 1)((ir + 6), ic)
                Next ic
            Next ir
            'rows 3-8
            For ir = 3 To 8
                For ic = 1 To 8
                    paFood(2, 3)(ir, ic) = paFood(1, 3)((ir - 2), ic)
                Next ic
            Next ir
        'Cube 4
            'rows 1-2 (last two rows from anim 1 cube 2)
            For ir = 1 To 2
                For ic = 1 To 8
                    paFood(2, 4)(ir, ic) = paFood(1, 2)((ir + 6), ic)
                Next ic
            Next ir
            'rows 3-8
            For ir = 3 To 8
                For ic = 1 To 8
                    paFood(2, 4)(ir, ic) = paFood(1, 4)((ir - 2), ic)
                Next ic
            Next ir
    
    'Anim: 3 (Points)
        'Cube 1
            'White pixels
                'rows 3-7
                For ir = 3 To 7
                    For ic = 3 To 5
                        paFood(3, 1)(ir, ic) = White
                    Next ic
                Next ir
                'row 5
                    paFood(3, 1)(5, 7) = White
                'rows 5-8
                For ir = 5 To 8
                    paFood(3, 1)(ir, 6) = White
                    paFood(3, 1)(ir, 8) = White
                Next ir
            'Black pixels
                'row 4,6
                    paFood(3, 1)(4, 4) = Black
                    paFood(3, 1)(4, 5) = Black
                    paFood(3, 1)(6, 3) = Black
                    paFood(3, 1)(6, 4) = Black
        'Cube 2
            'row 7
                paFood(3, 2)(7, 2) = White
            'rows 7-8
            For ir = 7 To 8
                paFood(3, 2)(ir, 1) = White
                paFood(3, 2)(ir, 3) = White
            Next ir
        'Cube 3
            'row 1
            For ic = 6 To 8
                paFood(3, 3)(1, ic) = White
            Next ic
        'Cube 4
            'row 1,3,5
                paFood(3, 4)(1, 5) = White
                paFood(3, 4)(3, 2) = White
                paFood(3, 4)(5, 5) = White
            'rows 1-3
            For ir = 1 To 3
                paFood(3, 4)(ir, 1) = White
                paFood(3, 4)(ir, 3) = White
            Next ir
            'rows 1-5
            For ir = 1 To 5
                paFood(3, 4)(ir, 4) = White
                paFood(3, 4)(ir, 6) = White
            Next ir

End Sub
