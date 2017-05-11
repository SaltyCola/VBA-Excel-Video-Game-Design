VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sprite_PacDots"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'=================================================================================='
'=================================== Properties ==================================='
'=================================================================================='

'PacDot Type Boolean
Private pPowerDot As Boolean 'True: this dot is a power dot

'PacDot Eaten Boolean
Private pEaten As Boolean 'True: this dot has been eaten, triggering powerup(PowerDot), or end of level(last PacDot)

'Coordinates
Private pXcol As Integer 'x, or column, coordinate of the top-left-most pixel
Private pYrow As Integer 'y, or row, coordinate of the top-left-most pixel

'Current Cubes for drawing
Private pCpac As Cube 'regular PacDot
Private pCpwr As Cube 'PowerDot
Private pCblnk As Cube 'PowerDot "Blink"

'Animation Number
Private pnAnim As Integer '1=regular PacDot, 2,3=PowerDot, 4=PowerDot "Blink"

'PacDot Color
Private pColor As Long 'color of the PacDot

'Pacdot Ranges
Private pColorRng As Range 'Range of the cells where the PacDot is drawn for recoloring during a non MsP collision

'''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
' Is this PacDot a PowerDot?
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Property Get PowerDot() As Boolean
    PowerDot = pPowerDot
End Property
Public Property Let PowerDot(Value As Boolean)
    pPowerDot = Value
End Property

'''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
' Has this PacDot been Eaten?
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
' Cube: PacDot
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Property Get Cpac() As Cube
    Set Cpac = pCpac
End Property
Public Property Let Cpac(Value As Cube)
    Set pCpac = Value
End Property

'''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
' Cube: PowerDot
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Property Get Cpwr() As Cube
    Set Cpwr = pCpwr
End Property
Public Property Let Cpwr(Value As Cube)
    Set pCpwr = Value
End Property

'''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
' Cube: PowerDot "Blink"
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Property Get Cblnk() As Cube
    Set Cblnk = pCblnk
End Property
Public Property Let Cblnk(Value As Cube)
    Set pCblnk = Value
End Property

'''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
' Animation Number
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
' PacDot Color
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Property Get Color() As Long
    Color = pColor
End Property
Public Property Let Color(Value As Long)
    pColor = Value
End Property

'''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
' PacDot Color Range
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Property Get ColorRng() As Range
    Set ColorRng = pColorRng
End Property
Public Property Let ColorRng(Value As Range)
    Set pColorRng = Value
End Property



'================================================================================='
'================================ Public Methods ================================='
'================================================================================='

Public Sub SetCubes()

    Dim ix As Integer
    Dim iy As Integer
    
    'Cube: PacDot
        'set position
            pCpac.Xcol = pXcol
            pCpac.Yrow = pYrow
            pCpac.SetPixels
        'color pixels
            pCpac.Pixels(4, 4).Color = pColor
            pCpac.Pixels(4, 5).Color = pColor
            pCpac.Pixels(5, 4).Color = pColor
            pCpac.Pixels(5, 5).Color = pColor

    'Cube: PowerDot
        'set position
            pCpwr.Xcol = pXcol
            pCpwr.Yrow = pYrow
            pCpwr.SetPixels
        'color pixels
            For iy = 1 To 8
                For ix = 3 To 6
                    pCpwr.Pixels(iy, ix).Color = pColor
                Next ix
            Next iy
            For iy = 3 To 6
                For ix = 1 To 8
                    pCpwr.Pixels(iy, ix).Color = pColor
                Next ix
            Next iy
            pCpwr.Pixels(2, 2).Color = pColor
            pCpwr.Pixels(2, 7).Color = pColor
            pCpwr.Pixels(7, 2).Color = pColor
            pCpwr.Pixels(7, 7).Color = pColor

    'Cube: PowerDot Blink
        'set position
            pCblnk.Xcol = pXcol
            pCblnk.Yrow = pYrow
            pCblnk.SetPixels
        'color pixels
            'pixels default to black

End Sub

Public Sub SetColorRange()

    Dim i As Integer 'iterator
    
    'Only need ranges for regular PacDots (not PowerDots, as those are redrawn every frame with the rest of the sprites)
    If Not pPowerDot Then
        Set pColorRng = Range(Cells(pYrow + 3, pXcol + 3), Cells(pYrow + 4, pXcol + 4))
    End If

End Sub

Public Sub Update()

    Dim i As Integer 'iterator
    
    'change animation number if a PowerDot
    If pPowerDot Then
        If pnAnim = 2 Then
            pnAnim = 3
        ElseIf pnAnim = 3 Then
            pnAnim = 4
        ElseIf pnAnim = 4 Then
            pnAnim = 2
        End If
    End If
    
    'fix pacdot color after non MsP collision
    If Not pPowerDot Then
        If pColorRng.Interior.Color <> pColor Then
            pColorRng.Interior.Color = pColor
        End If
    End If
    
    'collision with MsP
    If Not pEaten And (MsP.nDir = 1 And (MsP.Xcol + 4) = pXcol And (MsP.Yrow + 0) = pYrow) Or (MsP.nDir = 2 And (MsP.Xcol + 0) = pXcol And (MsP.Yrow + 4) = pYrow) Or (MsP.nDir = 3 And (MsP.Xcol + 8) = pXcol And (MsP.Yrow + 4) = pYrow) Or (MsP.nDir = 4 And (MsP.Xcol + 4) = pXcol And (MsP.Yrow + 8) = pYrow) Then
        'change pacdot to eaten
        pEaten = True
        'add dot's range to black color-range
        If Not pPowerDot Then
            Set CR_Array(1, 2) = Union(CR_Array(1, 2), pColorRng)
        End If
    End If
    
    'PacDot was Eaten
    If pEaten Then
        'update dots left in map
        Map.DotsLeft = Map.DotsLeft - 1
        'play sound
            PlayByteArray_Chomp
        'change Global Score Up date boolean
            ScoreUpdateBool = True
        'PacDot (10 pts)
        If Not pPowerDot Then
            'Update Scores
                Call UpdateScores(10)
        'PowerDot (50 pts)
        ElseIf pPowerDot Then
            'Make Ghosts Scared only if not in Eyes Mode
                If Blinky.GhostMode <> 4 Then: Blinky.GhostMode = 5 '5=Start Timer
                If Pinky.GhostMode <> 4 Then: Pinky.GhostMode = 5 '5=Start Timer
                If Inky.GhostMode <> 4 Then: Inky.GhostMode = 5 '5=Start Timer
                If Sue.GhostMode <> 4 Then: Sue.GhostMode = 5 '5=Start Timer
            'Ghost Back to Scared After Eyes
                If Blinky.GhostMode = 4 Then: Blinky.ScaredStateAgain = True 'another powerdot eaten
                If Pinky.GhostMode = 4 Then: Pinky.ScaredStateAgain = True 'another powerdot eaten
                If Inky.GhostMode = 4 Then: Inky.ScaredStateAgain = True 'another powerdot eaten
                If Sue.GhostMode = 4 Then: Sue.ScaredStateAgain = True 'another powerdot eaten
            'Update Scores
                Call UpdateScores(50)
        End If
    End If

End Sub

Public Sub Draw()

    'PacDot
    If pnAnim = 1 Then
        pCpac.Draw
    'PowerDot
    ElseIf pnAnim = 2 Or pnAnim = 3 Then
        pCpwr.Draw
    'PowerDot Blink
    ElseIf pnAnim = 4 Then
        pCblnk.Draw
    End If

End Sub


'=================================================================================='
'================================ Private Methods ================================='
'=================================================================================='

Private Sub Class_Initialize()

    Dim ix As Integer 'x iterator
    Dim iy As Integer 'y iterator
    
    'Initially not eaten
        pEaten = False
    
    'PacDot is initially not a PowerDot
        pPowerDot = False
        pnAnim = 1
    
    'Cube: PacDot
        Set pCpac = New Cube
    'Cube: PowerDot
        Set pCpwr = New Cube
    'Cube: PowerDot Blink
        Set pCblnk = New Cube

End Sub
