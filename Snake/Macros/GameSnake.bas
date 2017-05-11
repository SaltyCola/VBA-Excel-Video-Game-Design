Attribute VB_Name = "GameSnake"
'using kernel32 API
Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)


'public variables
Public GameInProg As Boolean 'Game in progress boolean
Public KeyPressed As Boolean 'Boolean to ensure only one key press per tick
Public Collision As Boolean 'true if snake head collides with border or snake body

Public Tick As Long 'Game Counter
Public ScorePoints As Long 'Score: Game Points
Public ScorePellets As Long 'Score: Pellets Consumed
Public xMin, xMax, yMin, yMax As Integer 'Center of edge cubes

Public ControlBoard As frm_Controls 'Control form
Public Snakey As Model_Snake 'Snake Model
Public Pellet As Model_Pellet 'Pellet Model

Public DirUp As Boolean 'for upward movement of head
Public DirLeft As Boolean 'for Leftward movement of head
Public DirRight As Boolean 'for Rightward movement of head
Public DirDown As Boolean 'for downward movement of head

Public Black As Long 'black color for snake body and pellets
Public White As Long 'white color for empty cubes
Public Red As Long 'red color for collision animation


Public Sub StartNewGame()
    
    'Clear Previous Game
    Call EndGame
    Range("A1").Select 'move selected cell out of the way

    'initialize public variables
    GameInProg = True
    KeyPressed = False
    Collision = False
    Tick = 1
    ScorePoints = 0
    ScorePellets = 0
    Set Snakey = New Model_Snake
    Set Pellet = New Model_Pellet
    xMin = 8 'Left for columns
    xMax = 305 'Right for columns
    yMin = 8 'Top for rows
    yMax = 239 'Bottom for rows
    Black = RGB(0, 0, 0)
    White = RGB(255, 255, 255)
    Red = RGB(255, 0, 0)
    DirUp = False
    DirLeft = False
    DirRight = False
    DirDown = False
    
    'clear screen
    Range("G7:KT240").Interior.Color = White
    
    'call Snakey initializer
    Call SnakeInitializer
    
    'create control form
    Set ControlBoard = New frm_Controls
    ControlBoard.Show vbModeless
    DoEvents
    
    'call game counter
    Call Counter

End Sub

Public Sub SnakeInitializer()

    Dim xUpperBound As Integer 'For random number generator
    Dim xLowerBound As Integer 'For random number generator
    Dim yUpperBound As Integer 'For random number generator
    Dim yLowerBound As Integer 'For random number generator
    Dim sDir As Integer 'For snake model initialization only

    'initialize scores and time
    ScorePoints = 0
    ActiveSheet.txtPoints.Value = 0
    ScorePellets = 0
    ActiveSheet.txtPellets.Value = 0
    ActiveSheet.txtLength.Value = 0
    ActiveSheet.txtTime.Value = "0.00 sec"

    'initialize bounds
    xLowerBound = 6
    xUpperBound = (300 / 3) - 5
    yLowerBound = 6
    yUpperBound = (234 / 3) - 5

    'randomly choose snake head position
        Randomize 'generate seed value for Rnd
    Snakey.Head.Xpos = 3 * Int((xUpperBound - xLowerBound + 1) * Rnd + xLowerBound)
        Randomize 'generate seed value for Rnd
    Snakey.Head.Ypos = 3 * Int((yUpperBound - yLowerBound + 1) * Rnd + yLowerBound)
    'randomly choose snake direction
        Randomize 'generate seed value for Rnd
    sDir = Int((4 - 1 + 1) * Rnd + 1) '1:up, 2:left, 3:right, 4:down
    If sDir = 1 Then 'Up
        DirUp = True
        DirLeft = False
        DirRight = False
        DirDown = False
    ElseIf sDir = 2 Then 'Left
        DirUp = False
        DirLeft = True
        DirRight = False
        DirDown = False
    ElseIf sDir = 3 Then 'Right
        DirUp = False
        DirLeft = False
        DirRight = True
        DirDown = False
    ElseIf sDir = 4 Then 'Down
        DirUp = False
        DirLeft = False
        DirRight = False
        DirDown = True
    End If
    
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'set length
    Snakey.Length = 0
    'add body section
    Snakey.AddBody
    'Initial Body Section Position (one cube away from head)
    If DirUp Then
        Snakey.BodySections.Item(1).Xpos = Snakey.Head.Xpos
        Snakey.BodySections.Item(1).Ypos = Snakey.Head.Ypos + 3
    ElseIf DirLeft Then
        Snakey.BodySections.Item(1).Xpos = Snakey.Head.Xpos + 3
        Snakey.BodySections.Item(1).Ypos = Snakey.Head.Ypos
    ElseIf DirRight Then
        Snakey.BodySections.Item(1).Xpos = Snakey.Head.Xpos - 3
        Snakey.BodySections.Item(1).Ypos = Snakey.Head.Ypos
    ElseIf DirDown Then
        Snakey.BodySections.Item(1).Xpos = Snakey.Head.Xpos
        Snakey.BodySections.Item(1).Ypos = Snakey.Head.Ypos - 3
    End If
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'Initial Tail Position (two cubes away from head)
    If DirUp Then
        Snakey.Tail.Xpos = Snakey.Head.Xpos
        Snakey.Tail.Ypos = Snakey.Head.Ypos + 6
    ElseIf DirLeft Then
        Snakey.Tail.Xpos = Snakey.Head.Xpos + 6
        Snakey.Tail.Ypos = Snakey.Head.Ypos
    ElseIf DirRight Then
        Snakey.Tail.Xpos = Snakey.Head.Xpos - 6
        Snakey.Tail.Ypos = Snakey.Head.Ypos
    ElseIf DirDown Then
        Snakey.Tail.Xpos = Snakey.Head.Xpos
        Snakey.Tail.Ypos = Snakey.Head.Ypos - 6
    End If
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

''''''''''''''''''''''''''''''''''''''''''''''''''''
    'Initial Eraser Position (two cubes away from head)
    If DirUp Then
        Snakey.Eraser.Xpos = Snakey.Head.Xpos
        Snakey.Eraser.Ypos = Snakey.Head.Ypos + 9
    ElseIf DirLeft Then
        Snakey.Eraser.Xpos = Snakey.Head.Xpos + 9
        Snakey.Eraser.Ypos = Snakey.Head.Ypos
    ElseIf DirRight Then
        Snakey.Eraser.Xpos = Snakey.Head.Xpos - 9
        Snakey.Eraser.Ypos = Snakey.Head.Ypos
    ElseIf DirDown Then
        Snakey.Eraser.Xpos = Snakey.Head.Xpos
        Snakey.Eraser.Ypos = Snakey.Head.Ypos - 9
    End If
''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    'Load Cubes (body section cubes are loaded in Model_Snake.AddBody method)
    Snakey.Head.SetCubes
    Snakey.Tail.SetCubes
    Snakey.Eraser.SetCubes
    Pellet.SetCubes
    
    'call Drawers
    Pellet.Draw
    Snakey.Draw
    
    'set velocity
    If DirUp Then
        Snakey.Head.Xvel = 0
        Snakey.Head.Yvel = -1
    ElseIf DirLeft Then
        Snakey.Head.Xvel = -1
        Snakey.Head.Yvel = 0
    ElseIf DirRight Then
        Snakey.Head.Xvel = 1
        Snakey.Head.Yvel = 0
    ElseIf DirDown Then
        Snakey.Head.Xvel = 0
        Snakey.Head.Yvel = 1
    End If

End Sub

Public Sub Counter()
    
    While GameInProg
        
        'tick every 3-100ths of a second
        Sleep 30 'wait 0.03 seconds
        
        'Run Game Counter in the background
        DoEvents
        
        'call Game Updater
        If GameInProg Then: Call GameUpdater
        
        'increment game counter and points
        Tick = Tick + 1
        ScorePoints = ScorePoints + 10
        
        'refresh KeyPressed boolean for next tick
        KeyPressed = False
        
    Wend

End Sub

Public Sub GameUpdater()

    Dim t3s As Long 'Tick in increments of 3
    Dim t100ths As Integer '100ths Digits of Tick counter
    Dim t10ths As Integer '10ths Digits of Tick counter
    Dim tInt As Integer 'Integer Digits of Tick counter
    Dim b As Integer 'integer for iterating backwards through the body sections collection
    Dim prevX As Long 'previous segment's X position before update
    Dim prevY As Long 'previous segment's Y position before update
    Dim cChkColl As Range 'range object for checking for Snake Collisions
    
    
    'Counter print out
    t3s = Tick * 3
    If t3s < 10 Then
        t100ths = t3s
        t10ths = 0
        tInt = 0
    ElseIf t3s < 100 Then
        t100ths = t3s Mod 10
        t10ths = Int((t3s - t100ths) / 10)
        tInt = 0
    Else
        t100ths = t3s Mod 10
        t10ths = Int(((t3s Mod 100) - t100ths) / 10)
        tInt = Int((t3s - (t10ths + t100ths)) / 100)
    End If
    ActiveSheet.txtTime.Text = tInt & "." & t10ths & t100ths & " sec"
    
    
    'Update Scores
    ActiveSheet.txtPoints.Text = ScorePoints
    ActiveSheet.txtPellets.Text = ScorePellets
    ActiveSheet.txtLength.Text = Snakey.Length + 2
    
    
    'Update Snake position (from eraser to head)
        'eraser
        prevX = Snakey.Tail.Xpos
        prevY = Snakey.Tail.Ypos
        Snakey.Eraser.Xpos = prevX
        Snakey.Eraser.Ypos = prevY
        'tail
        prevX = Snakey.BodySections.Item(Snakey.Length).Xpos
        prevY = Snakey.BodySections.Item(Snakey.Length).Ypos
        Snakey.Tail.Xpos = prevX
        Snakey.Tail.Ypos = prevY
        'body sections
        If Snakey.Length > 1 Then
            For b = Snakey.Length To 2 Step -1
                prevX = Snakey.BodySections.Item(b - 1).Xpos
                prevY = Snakey.BodySections.Item(b - 1).Ypos
                Snakey.BodySections.Item(b).Xpos = prevX
                Snakey.BodySections.Item(b).Ypos = prevY
            Next b
        End If
        prevX = Snakey.Head.Xpos
        prevY = Snakey.Head.Ypos
        Snakey.BodySections.Item(1).Xpos = prevX
        Snakey.BodySections.Item(1).Ypos = prevY
        'head
        Snakey.Head.Xpos = Snakey.Head.Xpos + (Snakey.Head.Xvel * 3)
        Snakey.Head.Ypos = Snakey.Head.Ypos + (Snakey.Head.Yvel * 3)
    
    
    'Environment Interactions
        Set cChkColl = Cells(yMin + Snakey.Head.Ypos, xMin + Snakey.Head.Xpos)
        'snake collision
        If (cChkColl.Interior.Color = Black) And Not ((Snakey.Head.Xpos = Pellet.Xpos) And (Snakey.Head.Ypos = Pellet.Ypos)) Then
            Collision = True
        'pellet eaten
        ElseIf (cChkColl.Interior.Color = Black) And ((Snakey.Head.Xpos = Pellet.Xpos) And (Snakey.Head.Ypos = Pellet.Ypos)) Then
            ScorePellets = ScorePellets + 1
            ScorePoints = ScorePoints + 5000
            Snakey.AddBody
            Set Pellet = New Model_Pellet
            Pellet.SetCubes
            Pellet.Draw
        End If
    
    
    'Draw Handler
        If Collision Then
            'stop game updater
            GameInProg = False
            'Call Collision Animator
            Snakey.CollisionAnimation
        Else
            'Draw Snake
            Snakey.Draw
        End If


End Sub

Public Sub EndGame()

    If GameInProg Or Collision Then
        Call frm_Controls.UserForm_Terminate
        ActiveSheet.lblGameOver.Visible = False
        ControlBoard.Hide
    End If

End Sub
