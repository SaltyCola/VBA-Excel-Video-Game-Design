Attribute VB_Name = "GameMsPacman"
'Code written by: Cody Normington
'Only 1st level playable so far, I plan to continue creating until the game is
' a full copy of the Original Arcade Game Ms. Pacman.

Option Explicit
'=====================================Declarations=======================================

'Sleep API sub to add time between frames (lower framerate)
Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

'Playing Multiple Sound Files at Once
Public Declare PtrSafe Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" _
    (ByVal lpstrCommand As String, ByVal lpstrReturnString As String, _
    ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long
Public Declare PtrSafe Function mciGetErrorString Lib "winmm.dll" Alias "mciGetErrorStringA" _
    (ByVal dwError As Long, ByVal lpstrBuffer As String, ByVal uLength As Long) As Long

'Loading / Play Audio from memory (byte arrays)
Public Const SND_SYNC = &H0
Public Const SND_ASYNC = &H1
Public Const SND_MEMORY = &H4
Public Const SND_FILENAME = &H20000
Const OPEN_EXISTING = 3
Const GENERIC_READ = &H80000000
Public Declare PtrSafe Function ReadFile Lib "kernel32" (ByVal hFile As Long, lpBuffer As Any, _
    ByVal nNumberOfBytesToRead As Long, lpNumberOfBytesToRead As Long, ByVal lpOverlapped As Any) As Long
Public Declare PtrSafe Function CreateFile Lib "kernel32" Alias "CreateFileA" _
    (ByVal lpFileName As String, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, _
    ByVal lpSecurityAttributes As Any, ByVal dwCreationDisposition As Long, _
    ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
Public Declare PtrSafe Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Public Declare PtrSafe Function PlaySound Lib "winmm.dll" Alias "PlaySoundA" (lpBuffer As Any, ByVal hModule As Long, ByVal dwFlags As Long) As Long

'===================================Public Variables=====================================

'game control variables
Public Wrkb As Workbook 'Game's workbook
Public Wrks As Worksheet 'Game's worksheet
Public GameInProgress As Boolean 'True: Game is currently running ; False: Game is not running
Public ControlForm As frm_Controls 'Reads keyboard input for directional comands
Public KeyState(1 To 8, 1 To 2) As Variant 'Array Holding the keystate of each of the 8 possible button presses (1:up, 2:left, 3:right, 4:down, 5:w, 6:a, 7:d, 8:s)
    '(For each of 8: 1=longCode, 2=ButtonPressedBoolean)
Public BoolKeyState As Boolean 'Turns off User Input during animations of Ms. Pacman such as tunneling

'Audio Arrays
Public bytearrChomp() As Byte 'Byte Array for wav file of similar name
Public arrExtraLife(1 To 2) As Variant '1=path ; 2 =playing boolean
Public arrFoodAmbient(1 To 2, 1 To 4) As Variant 'For both Food Ambient A and B, 1=path ; 2=playing boolean ; 3=repeat timer ; 4=reset timers boolean
Public arrFoodEat(1 To 2) As Variant '1=path ; 2 =playing boolean
Public arrGameStart(1 To 2) As Variant '1=path ; 2 =playing boolean
Public arrGhostAmbient(1 To 2, 1 To 4) As Variant 'For both Ghost Ambient A and B, 1=path ; 2=playing boolean ; 3=repeat timer ; 4=reset timers boolean
Public arrGhostEat(1 To 2) As Variant '1=path ; 2 =playing boolean
Public arrGhostEyes(1 To 2) As Variant '1=path ; 2 =playing boolean
Public arrAct1(1 To 2) As Variant '1=path ; 2 =playing boolean
Public arrAct2(1 To 2) As Variant '1=path ; 2 =playing boolean
Public arrAct3(1 To 2) As Variant '1=path ; 2 =playing boolean
Public arrMenuSelect(1 To 2) As Variant '1=path ; 2 =playing boolean
Public arrMsPDeath(1 To 2) As Variant '1=path ; 2 =playing boolean
Public arrPowerDotAmbient(1 To 2) As Variant '1=path ; 2 =playing boolean (different from other ambients as it has a definitely cut-off point)

'color variables
Public Black As Long 'Color-Range Array: 1 (allows reset of all moving sprites)
Public White As Long 'Color-Range Array: 2
Public Yellow As Long 'Color-Range Array: 3
Public Red As Long 'Color-Range Array: 4
Public Blue As Long 'Color-Range Array: 5
Public Pink As Long 'Color-Range Array: 6
Public Cyan As Long 'Color-Range Array: 7
Public Orange As Long 'Color-Range Array: 8
Public Brown As Long 'Color-Range Array: 9
Public Green As Long 'Color-Range Array: 10
Public Salmon As Long 'Not a color-range
Public Grey As Long 'Not a color-range
Public LightBlue As Long 'Not a color-range
Public TrackBlack As Long 'Not a color-range

'Level Map Object
Public Map As Sprite_Map 'Holds information required to generate and load levels
Public LvlNum As Integer 'Level Number
Public PwrDotDuration As Integer 'Holds the current level's switch timer duration
Public PwrDotNonFlash As Integer 'Holds the current level's time until flashing segment begins

'Tunnel and Center Corner Locations
Public NumberOfTunnels As Integer 'Number of tunnels will tell food ai whether to use 2 or 4 for random choosing
Public TunnelCoords() As Integer 'Array of tunnel coords (1:[x,y],2:...)
Public CenterCornerCoords() As Integer 'Array of central corner coords (1:[x,y],2:...)

'Scoring and Text
Public HeaderRange As Range 'Range of scoring text objects for blacking out before update
Public ScoreUpdateBool As Boolean 'True: Score has been updated
Public txtHighScore As GameText 'text object "High Score"
Public scoreHighScore As GameText 'text object score that appears under "High Score"
Public valueHighScore As Long 'score value to be given to text object
Public txt1Up As GameText 'text object "1Up"
Public score1Up As GameText 'text object score that appears under "1UP"
Public value1Up As Long 'score value to be given to text object

'Center Map Text
Public txtReady As GameText 'Ready! message before starting a round
Public txtGameOver As GameText 'Game Over message for dying on last life

'Animations
Public GameStartAnim As Integer '0=regular gameplay ; 1=Game Start Animation is running ; 2=Allows first frame of sprites

'Color-Range Array
Public CR_Array(1 To 10, 1 To 2) As Variant 'each of 10: 1st pos = color long, 2nd pos = range
    Public CLR_Array(1 To 9, 1 To 2) As Variant 'array to assist in speeding up frame loading: 1=counter, 2=range

'Map-Range Array
Public MR_Array(1 To 6, 1 To 2) As Variant 'each of 6: 1st pos = color long, 2nd pos = range
    'Map Loading Range Array
    Public MLR_Array(1 To 7, 1 To 2) As Variant 'array to assist in speeding up map loading: 1=counter, 2=range

'Sprite Objects
Public MsP As Sprite_MsPacman 'Ms. Pacman's Sprite object
    Public MsPLives As Integer 'Number of Extra Lives MsP has left (0 -> last life)
    Public MsPLivesArr() As Sprite_MsPacman 'Array Holding MsP's extra lives sprite objects
Public Blinky As Sprite_Ghosts 'Blinky's Sprite object
Public Pinky As Sprite_Ghosts 'Pinky's Sprite object
Public Inky As Sprite_Ghosts 'Inky's Sprite object
Public Sue As Sprite_Ghosts 'Sue's Sprite object
    Public GhostPhaseType As Integer '1=Chase ; 2=Scatter
    Public GhostTimerChase As Integer 'Times the Chase phase of the ghosts
    Public GhostTimerScatter As Integer 'Times the Scatter phase of the ghosts
    Public GhostScatterNumber As Integer 'Number of scatter phases per level or per life (reset on new lives and new levels)
Public Food As Sprite_Food 'Sprite object for all food bonuses
    Public FoodTypesArr() As Sprite_Food 'Array Holding Food Sprites for map footer
    Public FoodTypesNumber As Integer 'Tells array how many to draw during each level

'Sound File Variables
Public glo_from As Long 'first frame of audio file to start playing at
Public glo_to As Long 'last frame of audio file to stop playing at
Public glo_AliasName As String 'filename of audio file
Public glo_hWnd As Long '???

'======================================Subroutines========================================

Public Sub MsPacman()

    'Turn On Optimizers
'    Application.Calculation = xlCalculationManual
    
    'initialize worksheet variables
    Set Wrkb = Workbooks("MsP")
    Set Wrks = Wrkb.Worksheets("MsP")
    
    'activate worksheet if not already
    Wrks.Activate
    
    'black out screen
    Wrks.Rows.Interior.Color = RGB(0, 0, 0) 'clear screen
    
    'initialize application window
    Application.ExecuteExcel4Macro "Show.ToolBar(""Ribbon"", False)" 'hide ribbon
    Application.WindowState = xlNormal 'unmaximize window
    Application.top = 0 'set window position
    Application.left = 0 'set window position
    Application.Height = 618 'set window size
    Application.Width = 434 'set window size
    
    'select cell "A1"
    Wrks.Cells(1, 1).Select
    
    'Initialize Game
    Call Initialize

End Sub
    
Public Sub Initialize()

    Dim i As Integer 'integer iterator
    Dim i1 As Integer 'integer iterator
    
    'set initial key state
        BoolKeyState = True
        KeyState(1, 1) = vbKeyUp
        KeyState(2, 1) = vbKeyLeft
        KeyState(3, 1) = vbKeyRight
        KeyState(4, 1) = vbKeyDown
        KeyState(5, 1) = vbKeyW
        KeyState(6, 1) = vbKeyA
        KeyState(7, 1) = vbKeyD
        KeyState(8, 1) = vbKeyS
        For i = 1 To 8
            KeyState(i, 2) = False 'no keys pressed yet
        Next i
    
    'initialize control form
        Set ControlForm = New frm_Controls
        
    'initialize audio file arrays
        LoadWaveIntoMemory Wrkb.Path & "\Audio\wav16bit\MsP_Chomp.wav"
        arrExtraLife(1) = Wrkb.Path & "\Audio\wav16bit\MsP_ExtraLife.wav"
            arrExtraLife(2) = False 'playing boolean
        arrFoodAmbient(1, 1) = Wrkb.Path & "\Audio\wav16bit\MsP_FoodAmbientA.wav"
            arrFoodAmbient(1, 2) = False 'playing boolean
            arrFoodAmbient(1, 3) = 2000 'counter for switch between A and B
            arrFoodAmbient(1, 4) = True 'boolean for switch between A and B
        arrFoodAmbient(2, 1) = Wrkb.Path & "\Audio\wav16bit\MsP_FoodAmbientB.wav"
            arrFoodAmbient(2, 2) = False 'playing boolean
            arrFoodAmbient(2, 3) = 0 'counter for switch between A and B
            arrFoodAmbient(2, 4) = True 'boolean for switch between A and B
        arrFoodEat(1) = Wrkb.Path & "\Audio\wav16bit\MsP_FoodEat.wav"
            arrFoodEat(2) = False  'playing boolean
        arrGameStart(1) = Wrkb.Path & "\Audio\wav16bit\MsP_GameStart.wav"
            arrGameStart(2) = False  'playing boolean
        arrGhostAmbient(1, 1) = Wrkb.Path & "\Audio\wav16bit\MsP_GhostAmbientA.wav"
            arrGhostAmbient(1, 2) = False  'playing boolean
            arrGhostAmbient(1, 3) = 2000 'counter for switch between A and B
            arrGhostAmbient(1, 4) = True 'boolean for switch between A and B
        arrGhostAmbient(2, 1) = Wrkb.Path & "\Audio\wav16bit\MsP_GhostAmbientB.wav"
            arrGhostAmbient(2, 2) = False  'playing boolean
            arrGhostAmbient(2, 3) = 0 'counter for switch between A and B
            arrGhostAmbient(2, 4) = True 'boolean for switch between A and B
        arrGhostEat(1) = Wrkb.Path & "\Audio\wav16bit\MsP_GhostEat.wav"
            arrGhostEat(2) = False  'playing boolean
        arrGhostEyes(1) = Wrkb.Path & "\Audio\wav16bit\MsP_GhostEyes.wav"
            arrGhostEyes(2) = False  'playing boolean
        arrAct1(1) = Wrkb.Path & "\Audio\wav16bit\MsP_IntermissionAct1_TheyMeet.wav"
            arrAct1(2) = False  'playing boolean
        arrAct2(1) = Wrkb.Path & "\Audio\wav16bit\MsP_IntermissionAct2_TheChase.wav"
            arrAct2(2) = False  'playing boolean
        arrAct3(1) = Wrkb.Path & "\Audio\wav16bit\MsP_IntermissionAct3_Junior.wav"
            arrAct3(2) = False  'playing boolean
        arrMenuSelect(1) = Wrkb.Path & "\Audio\wav16bit\MsP_MenuSelect.wav"
            arrMenuSelect(2) = False  'playing boolean
        arrMsPDeath(1) = Wrkb.Path & "\Audio\wav16bit\MsP_MsPDeath.wav"
            arrMsPDeath(2) = False  'playing boolean
        arrPowerDotAmbient(1) = Wrkb.Path & "\Audio\wav16bit\MsP_PowerDotAmbient.wav"
            arrPowerDotAmbient(2) = False  'playing boolean
    
    'set color variables
        Black = RGB(0, 0, 0)
        White = RGB(255, 255, 255)
        Yellow = RGB(255, 255, 0)
        Red = RGB(255, 0, 0)
        Blue = RGB(0, 0, 255)
        Pink = RGB(255, 100, 150)
        Cyan = RGB(0, 255, 255)
        Orange = RGB(255, 150, 0)
        Brown = RGB(255, 100, 0)
        Green = RGB(0, 255, 0)
        Salmon = RGB(255, 150, 100) 'MAP COLOR ONLY
        Grey = RGB(200, 200, 200) 'MAP COLOR ONLY
        LightBlue = RGB(50, 150, 255) 'MAP COLOR ONLY
        TrackBlack = RGB(3, 3, 3) 'MAP COLOR ONLY
        
    'initialize Scoring Header Range
        Set HeaderRange = Range(Cells(9, 1), Cells(24, 256)) 'allows for updating scores by blacking out the previous score
        ScoreUpdateBool = True
    
    'initialize Range Arrays
        'CR colors
        CR_Array(1, 1) = Black
        CR_Array(2, 1) = White
        CR_Array(3, 1) = Yellow
        CR_Array(4, 1) = Red
        CR_Array(5, 1) = Blue
        CR_Array(6, 1) = Pink
        CR_Array(7, 1) = Cyan
        CR_Array(8, 1) = Orange
        CR_Array(9, 1) = Brown
        CR_Array(10, 1) = Green
        'CR ranges
        Set CR_Array(1, 2) = Union(Range(Cells(273, 1), Cells(288, 128)), Range(Cells(1, 1), Cells(288, 16)), Range(Cells(1, 241), Cells(288, 256))) 'black areas beneath(MsPLives only) and to the left and right of level map
        Set CR_Array(2, 2) = Cells(1, 2)
        Set CR_Array(3, 2) = Cells(1, 3)
        Set CR_Array(4, 2) = Cells(1, 4)
        Set CR_Array(5, 2) = Cells(1, 5)
        Set CR_Array(6, 2) = Cells(1, 6)
        Set CR_Array(7, 2) = Cells(1, 7)
        Set CR_Array(8, 2) = Cells(1, 8)
        Set CR_Array(9, 2) = Cells(1, 9)
        Set CR_Array(10, 2) = Cells(1, 10)
        'MR ranges
        Set MR_Array(1, 2) = Cells(1, 11)
        Set MR_Array(2, 2) = Cells(1, 12)
        Set MR_Array(3, 2) = Cells(1, 13)
        Set MR_Array(4, 2) = Cells(1, 14)
        Set MR_Array(5, 2) = Cells(1, 15)
        Set MR_Array(6, 2) = Cells(1, 16)
        'MLR counters
        MLR_Array(1, 1) = 0
        MLR_Array(2, 1) = 0
        MLR_Array(3, 1) = 0
        MLR_Array(4, 1) = 0
        MLR_Array(5, 1) = 0
        MLR_Array(6, 1) = 0
        MLR_Array(7, 1) = 0
        'MLR ranges
        Set MLR_Array(1, 2) = Cells(2, 1) 'corresponds to MR_Array(1,2)
        Set MLR_Array(2, 2) = Cells(2, 2) 'corresponds to MR_Array(2,2)
        Set MLR_Array(3, 2) = Cells(2, 3) 'corresponds to MR_Array(3,2)
        Set MLR_Array(4, 2) = Cells(2, 4) 'corresponds to MR_Array(4,2)
        Set MLR_Array(5, 2) = Cells(2, 5) 'corresponds to MR_Array(5,2)
        Set MLR_Array(6, 2) = Cells(2, 6) 'corresponds to MR_Array(6,2)
        Set MLR_Array(7, 2) = Cells(2, 7) 'corresponds to CR_Array(1,2)
        'CLR counters
        CLR_Array(1, 1) = 0
        CLR_Array(2, 1) = 0
        CLR_Array(3, 1) = 0
        CLR_Array(4, 1) = 0
        CLR_Array(5, 1) = 0
        CLR_Array(6, 1) = 0
        CLR_Array(7, 1) = 0
        CLR_Array(8, 1) = 0
        CLR_Array(9, 1) = 0
        'CLR ranges
        Set CLR_Array(1, 2) = Cells(2, 8) 'corresponds to CR_Array(2,2)
        Set CLR_Array(2, 2) = Cells(2, 9) 'corresponds to CR_Array(3,2)
        Set CLR_Array(3, 2) = Cells(2, 10) 'corresponds to CR_Array(4,2)
        Set CLR_Array(4, 2) = Cells(2, 11) 'corresponds to CR_Array(5,2)
        Set CLR_Array(5, 2) = Cells(2, 12) 'corresponds to CR_Array(6,2)
        Set CLR_Array(6, 2) = Cells(2, 13) 'corresponds to CR_Array(7,2)
        Set CLR_Array(7, 2) = Cells(2, 14) 'corresponds to CR_Array(8,2)
        Set CLR_Array(8, 2) = Cells(2, 15) 'corresponds to CR_Array(9,2)
        Set CLR_Array(9, 2) = Cells(2, 16) 'corresponds to CR_Array(10,2)
        
        'initialize hidden sides for use of level tunnels
        For i = 1 To 31 'total number of cubes
            For i1 = 1 To 16 'left two columns
                Set MR_Array(3, 2) = Union(MR_Array(3, 2), Cells(((i * 8) + 20), (i1)))
                Set MR_Array(3, 2) = Union(MR_Array(3, 2), Cells((((i * 8) + 20) + 1), (i1)))
            Next i1
            For i1 = 1 To 16 'right two columns
                Set MR_Array(3, 2) = Union(MR_Array(3, 2), Cells(((i * 8) + 20), (i1 + 240)))
                Set MR_Array(3, 2) = Union(MR_Array(3, 2), Cells((((i * 8) + 20) + 1), (i1 + 240)))
            Next i1
        Next i
    
    'Initialize Ghosts
        'Blinky
            Set Blinky = New Sprite_Ghosts
            Blinky.GhostName = "Blinky"
            'set GhostCage coords
            Blinky.XcolGC = 121
            Blinky.YrowGC = 133
            'set ghost coords
            Blinky.Xcol = 121
            Blinky.Yrow = 109
            Blinky.SetPosition
            'set ai move type
            Blinky.AIModule.MoveType = 0
            Blinky.AIModule.CageMoveType = 0
            Blinky.AIModule.CageMoveCounter = 0
            'set ghost color
            Blinky.GhostColor = Red
            'set initial direction
            Blinky.nDir = 2
            'create sprites
            Blinky.InitializeSprites
        'Pinky
            Set Pinky = New Sprite_Ghosts
            Pinky.GhostName = "Pinky"
            'set GhostCage coords
            Pinky.XcolGC = 121
            Pinky.YrowGC = 133
            'set ghost coords
            Pinky.Xcol = 121
            Pinky.Yrow = 133
            Pinky.SetPosition
            'set ai move type
            Pinky.AIModule.MoveType = 0
            Pinky.AIModule.CageMoveType = 2
            Pinky.AIModule.CageMoveCounter = 0
            'set ghost color
            Pinky.GhostColor = Pink
            'set initial direction
            Pinky.nDir = 4
            'create sprites
            Pinky.InitializeSprites
        'Inky
            Set Inky = New Sprite_Ghosts
            Inky.GhostName = "Inky"
            'set GhostCage coords
            Inky.XcolGC = 105
            Inky.YrowGC = 133
            'set ghost coords
            Inky.Xcol = 105
            Inky.Yrow = 133
            Inky.SetPosition
            'set ai move type
            Inky.AIModule.MoveType = 0
            Inky.AIModule.CageMoveType = 2
            Inky.AIModule.CageMoveCounter = 0
            'set ghost color
            Inky.GhostColor = Cyan
            'set initial direction
            Inky.nDir = 1
            'create sprites
            Inky.InitializeSprites
        'Sue
            Set Sue = New Sprite_Ghosts
            Sue.GhostName = "Sue"
            'set GhostCage coords
            Sue.XcolGC = 137
            Sue.YrowGC = 133
            'set ghost coords
            Sue.Xcol = 137
            Sue.Yrow = 133
            Sue.SetPosition
            'set ai move type
            Sue.AIModule.MoveType = 0
            Sue.AIModule.CageMoveType = 2
            Sue.AIModule.CageMoveCounter = 0
            'set ghost color
            Sue.GhostColor = Orange
            'set initial direction
            Sue.nDir = 1
            'create sprites
            Sue.InitializeSprites
    
    'intialize map object
        Set Map = New Sprite_Map
        'set map dimensions
        Map.WidthCubes = 28
        Map.HeightCubes = 31
        'load level maps
        Call LoadLevel(1)
    
    'initialize MsP
        Set MsP = New Sprite_MsPacman
        'set MsP coords
        MsP.Xcol = 121 + 4 'Starting at coordinate of the 14th cube across (+ 4 for first frame movement before game start animation)
        MsP.Yrow = 205 'Starting at a half-cube below 25th cube down
        MsP.SetPosition
        'create sprites
        MsP.InitializeSprites
        
    'initialize MsP Lives
        MsPLives = 4
        ReDim MsPLivesArr(1 To MsPLives) As Sprite_MsPacman
        For i = 2 To MsPLives '1st entry is MsP herself, so no extra object
            Set MsPLivesArr(i) = New Sprite_MsPacman
            With MsPLivesArr(i)
                .Xcol = 21 + (((i - 1) - 1) * 16) 'first i-1 is to ignore empty MsPLivesArr(1)
                .Yrow = 273
                .nAnim = 2 'Resting Anim
                .nDir = 3 'Facing Right
                .Vx = 0 'no velocity
                .Vy = 0 'no velocity
                .Speed = 0 'zero speed
                .SetPosition
                .InitializeSprites
                .SetAnimDir
            End With
        Next i
    
    'set game in prog bool
        GameInProgress = True
        
    'show control form
        ControlForm.Show vbModeless
        DoEvents
    
    'Start Game
        Call Start

End Sub

Public Sub Start()
    
    'game in progress loop
    Do Until Not GameInProgress
        DoEvents
'TESTING FRAMERATE<==========================================
'StartMT = MicroTimer * 60 'yields values in ticks (i.e. 1/60th of a second) want to get to a value of 1 tick per frame
'TESTING FRAMERATE<==========================================
        'Call Sprite Movement Subs Chain
        Call UpdatePositions
'TESTING FRAMERATE<==========================================
'EndMT = MicroTimer * 60 'yields values in ticks (i.e. 1/60th of a second) want to get to a value of 1 tick per frame
'MsgBox (EndMT - StartMT)
'TESTING FRAMERATE<==========================================
    Loop
    
    'Clear game board
    Call GameBoardClear

End Sub

Public Sub GameEnd()

    'only end if game is running
    If GameInProgress Then
        'end all audio
        Call EndAudio(arrExtraLife, False)
        Call EndAudio(arrFoodAmbient, True, "A")
        Call EndAudio(arrFoodAmbient, True, "B")
        Call EndAudio(arrFoodEat, False)
        Call EndAudio(arrGameStart, False)
        Call EndAudio(arrGhostAmbient, True, "A")
        Call EndAudio(arrGhostAmbient, True, "B")
        Call EndAudio(arrGhostEat, False)
        Call EndAudio(arrGhostEyes, False)
        Call EndAudio(arrAct1, False)
        Call EndAudio(arrAct2, False)
        Call EndAudio(arrAct3, False)
        Call EndAudio(arrMenuSelect, False)
        Call EndAudio(arrMsPDeath, False)
        Call EndAudio(arrPowerDotAmbient, False)
        'change boolean
        GameInProgress = False
        'hide ControlForm
        ControlForm.Hide
    End If

End Sub

Public Sub GameBoardClear()
    
    'black out screen
    Wrks.Rows.Interior.Color = Black 'clear screen
    
    'Turn off Optimizers
'    Application.Calculation = xlCalculationAutomatic
    
    'reset application window
    Application.ExecuteExcel4Macro "Show.ToolBar(""Ribbon"", True)" 'show ribbon
    Application.WindowState = xlMaximized 'maximize window
    
    'Erase Control Form
    Set ControlForm = Nothing

End Sub

Public Sub UpdatePositions()

    Dim i As Integer 'iterator
    
    'call Clearer
    Call ClearFrame
    
    'update positions:
    If GameStartAnim <> 1 Then
        
        'Set GameStartAnim to 1 after first frame of Sprites is drawn
            If GameStartAnim = 2 Then: GameStartAnim = 1
        
        'MsP Update
            MsP.UpdatePosition
        
        'Food
            Food.UpdatePosition
        
        'Ghosts
            'Blinky
            Blinky.UpdatePosition
            'Pinky
            Pinky.UpdatePosition
            'Inky
            Inky.UpdatePosition
            'Sue
            Sue.UpdatePosition
        
        'Ghost Phases
            If ((Blinky.AIModule.MoveType = 1 Or Blinky.AIModule.MoveType = 2) _
            Or (Pinky.AIModule.MoveType = 1 Or Pinky.AIModule.MoveType = 2) _
            Or (Inky.AIModule.MoveType = 1 Or Inky.AIModule.MoveType = 2) _
            Or (Sue.AIModule.MoveType = 1 Or Sue.AIModule.MoveType = 2)) _
            And (GhostScatterNumber > 0) Then
                If MsP.MsPMode = 1 Then
                    'decrement timers
                        'chase
                        If GhostPhaseType = 1 And GhostTimerChase > 0 Then
                            GhostTimerChase = GhostTimerChase - 1
                        'scatter
                        ElseIf GhostPhaseType = 2 And GhostTimerScatter > 0 Then
                            GhostTimerScatter = GhostTimerScatter - 1
                        End If
                    'switch ghost phases
                        'Chase to Scatter
                        If GhostPhaseType = 1 And GhostTimerChase = 0 Then
                            'Ghost Phase Change
                            Blinky.SwitchChaseScatter
                            Pinky.SwitchChaseScatter
                            Inky.SwitchChaseScatter
                            Sue.SwitchChaseScatter
                            'Globals
                            GhostPhaseType = 2
                            Call ResetGhostTimers(LvlNum)
                        'Scatter to Chase
                        ElseIf GhostPhaseType = 2 And GhostTimerScatter = 0 Then
                            'Ghost Phase Change
                            Blinky.SwitchChaseScatter
                            Pinky.SwitchChaseScatter
                            Inky.SwitchChaseScatter
                            Sue.SwitchChaseScatter
                            'Globals
                            GhostPhaseType = 1
                            GhostScatterNumber = GhostScatterNumber - 1
                            Call ResetGhostTimers(LvlNum)
                        End If
                End If
            End If
        
        'PacDots (list in Map Object)
            For i = 1 To Map.DotsCnt
                If Not Map.Dots(i).Eaten Then
                    Map.Dots(i).Update
                End If
            Next i
        
        'End Level
            If Map.DotsLeft = 0 Then
                Call GameEnd
            End If
    
    End If
    
    'call Drawer
    Call DrawFrame

End Sub

Public Sub ClearFrame()

    Dim i As Integer 'iterator
    
    'reset all color-ranges (except black: first entry)
        For i = 2 To 10
            Set CR_Array(i, 2) = Cells(1, i)
        Next i
    
    'reset all color-loading counters and ranges
        For i = 1 To 9
            CLR_Array(i, 1) = 0
            Set CLR_Array(i, 2) = Cells(2, (i + 7))
        Next i
        

End Sub

Public Sub DrawFrame()

    Dim i As Integer 'integer iterator
    
    'Draw Sprites:
        'PacDots (from Map)
            For i = 1 To Map.DotsCnt
                If Not Map.Dots(i).Eaten And Map.Dots(i).PowerDot Then
                    Map.Dots(i).Draw
                End If
            Next i
        'MsP
            MsP.Draw
        'Draw MsP Lives
            For i = 2 To MsPLives '1st entry is nothing (represents MsP herself)
                MsPLivesArr(i).Draw
            Next i
        'Food
            Food.Draw
        'Ghosts
            'Blinky
            Blinky.Draw
            'Pinky
            Pinky.Draw
            'Inky
            Inky.Draw
            'Sue
            Sue.Draw
    
    'Draw Game Start Animation and play Music
    If GameStartAnim = 1 Then
        txtReady.DrawText 'Ready! message
        'play game start music
        Call PlayAudio(arrGameStart, False)
    End If
    
    'Grab remaining ranges from loading array ranges
        For i = 1 To 9
            'dump remainder into cr_array
            Set CR_Array((i + 1), 2) = Union(CR_Array((i + 1), 2), CLR_Array(i, 2))
            'clear clr_array
            Set CLR_Array(i, 2) = Cells(2, (i + 7))
            CLR_Array(i, 1) = 0
        Next i
            
    'apply frame (black is applied first to clear previous frame)
    Application.ScreenUpdating = False
    'Score Header Range
    If ScoreUpdateBool Then
        HeaderRange.Interior.Color = Black
        ScoreUpdateBool = False
    End If
    'Color Ranges
    For i = 1 To UBound(CR_Array)
        If CR_Array(i, 2).Count <> 1 Then
            CR_Array(i, 2).Interior.Color = CR_Array(i, 1) 'color ranges (track range is also a part of black range)
        End If
    Next i
    'Ghost Cage Gate Fix
    Range("DR126:EE127").Interior.Color = Map.ColorGate
    Application.ScreenUpdating = True
    
    'Game Start Animation Pause
    If GameStartAnim = 1 Then
        'wait until end of music
        Application.Wait DateAdd("s", 5, Now)
        'End animation
        GameStartAnim = 0
        'Set Ghost Motion Properties
            'Phase Timers
            Call ResetGhostTimers(LvlNum)
            GhostPhaseType = 2
            'Blinky
            Blinky.Vx = -1
            Blinky.Vy = 0
            Blinky.AIModule.MoveType = 2
            Blinky.AIModule.TimerCageReset "Blinky"
            'Pinky
            Pinky.Vx = 0
            Pinky.Vy = 0
            Pinky.AIModule.MoveType = 0
            Pinky.AIModule.TimerCageReset "Pinky"
            'Inky
            Inky.Vx = 0
            Inky.Vy = 0
            Inky.AIModule.MoveType = 0
            Inky.AIModule.TimerCageReset "Inky"
            'Sue
            Sue.Vx = 0
            Sue.Vy = 0
            Sue.AIModule.MoveType = 0
            Sue.AIModule.TimerCageReset "Sue"
    End If
    
    'Call Audio Controller
    Call AmbientAudioController

End Sub

Public Sub ResetGhostTimers(ByVal levelNum As Integer)
'Resets both chase and scatter timers for Ghosts

    'Level 1
    If levelNum = 1 Then
        GhostTimerChase = 200 '200fs ~ 20sec
        GhostTimerScatter = 70 '70fs ~ 7sec
    End If

End Sub

Public Sub ResetGhostPhaseCounter(ByVal levelNum As Integer)
'Resets scatter number on MsP Death

    'Level 1
    If levelNum = 1 Then
        GhostScatterNumber = 4
    End If

End Sub

Public Sub LoadLevel(ByVal levelNum As Integer)

    Dim strLevelSeed As String 'level string to be given to map's Seed_Load method
    Dim i As Integer 'iterator
    Dim c As Range 'iterator
    
    'Set Public Level Number variable
        LvlNum = levelNum
    
    'Set Game Start Animation Boolean
        GameStartAnim = 2 'Allows for first frame of sprites to be drawn before animation
    
    'initialize Ghost Scared State Timer and Scatter/Chase Timers
        If LvlNum = 1 Then 'Level 1
            '100 frames overall, with 75 frames before flashing
            PwrDotDuration = 100
            PwrDotNonFlash = 75
            'Chase/Scatter Timers
            GhostPhaseType = 2 'starts in scatter phase
            GhostTimerChase = 200 '200fs ~ 20 seconds
            GhostTimerScatter = 70 '70fs ~ 7 seconds
            GhostScatterNumber = 4 'scatter phases cease after 4 (reset on new lives and new levels)
        End If
    
    'initialize Food
        If LvlNum = 1 Then 'Level 1
            'initialize object
            Set Food = New Sprite_Food
            'create sprites
            Food.InitializeSprites ("cherry")
        End If
    
    'Initialize Tunnel and Central Corner Locations
        If LvlNum = 1 Then 'Level 1
            'Number of Tunnels
            NumberOfTunnels = 4
            'Tunnels
            ReDim TunnelCoords(1 To NumberOfTunnels, 1 To 2) As Integer
            TunnelCoords(1, 1) = 1 'top-left x
            TunnelCoords(1, 2) = 85 'top-left y
            TunnelCoords(2, 1) = 241 'top-right x
            TunnelCoords(2, 2) = 85 'top-right y
            TunnelCoords(3, 1) = 1 'bottom-left x
            TunnelCoords(3, 2) = 157 'bottom-left y
            TunnelCoords(4, 1) = 241 'bottom-right x
            TunnelCoords(4, 2) = 157 'bottom-right y
            'Center Corners
            ReDim CenterCornerCoords(1 To 4, 1 To 2) As Integer
            CenterCornerCoords(1, 1) = 85 'top-left x
            CenterCornerCoords(1, 2) = 109 'top-left y
            CenterCornerCoords(2, 1) = 157 'top-right x
            CenterCornerCoords(2, 2) = 109 'top-right y
            CenterCornerCoords(3, 1) = 85 'bottom-left x
            CenterCornerCoords(3, 2) = 157 'bottom-left y
            CenterCornerCoords(4, 1) = 157 'bottom-right x
            CenterCornerCoords(4, 2) = 157 'bottom-right y
        End If
    
    'Initialize Food Types Array
        ReDim FoodTypesArr(1 To 7) As Sprite_Food
        For i = 1 To 7
            'initialize Food objects
            Set FoodTypesArr(i) = New Sprite_Food
            'create sprites
            If i = 1 Then: FoodTypesArr(i).InitializeSprites ("cherry")
            If i = 2 Then: FoodTypesArr(i).InitializeSprites ("strawberry")
            If i = 3 Then: FoodTypesArr(i).InitializeSprites ("peach")
            If i = 4 Then: FoodTypesArr(i).InitializeSprites ("pretzel")
            If i = 5 Then: FoodTypesArr(i).InitializeSprites ("apple")
            If i = 6 Then: FoodTypesArr(i).InitializeSprites ("pear")
            If i = 7 Then: FoodTypesArr(i).InitializeSprites ("banana")
            'set position
            With FoodTypesArr(i)
                .Xcol = 225 - ((i - 1) * 16)
                .Yrow = 273
                .SetPosition
            End With
        Next i
    
    'Set Food Types Number by level
        If LvlNum = 1 Then
            FoodTypesNumber = 1
        End If
    
    'Set Ghost Switch Timer Variables
        'Blinky
            Blinky.TimerDuration = PwrDotDuration
            Blinky.SwitchTimerMode = 2
            Blinky.SwitchTimer = PwrDotNonFlash
        'Pinky
            Pinky.TimerDuration = PwrDotDuration
            Pinky.SwitchTimerMode = 2
            Pinky.SwitchTimer = PwrDotNonFlash
        'Inky
            Inky.TimerDuration = PwrDotDuration
            Inky.SwitchTimerMode = 2
            Inky.SwitchTimer = PwrDotNonFlash
        'Sue
            Sue.TimerDuration = PwrDotDuration
            Sue.SwitchTimerMode = 2
            Sue.SwitchTimer = PwrDotNonFlash
    
    'initialize game score values
        If LvlNum = 1 Then
            valueHighScore = 0
            value1Up = 0
        End If
    
    'initialize text objects in header
        'High Score
        Set txtHighScore = New GameText
            txtHighScore.InitializePixelColor White
            txtHighScore.GenerateText "High Score", 1, 89
        Set scoreHighScore = New GameText
            scoreHighScore.SwitchJustification False, True 'Right Justified
            'reset high score
            scoreHighScore.InitializePixelColor White
            scoreHighScore.GenerateText "00", 9, 161
        '1Up Score
        Set txt1Up = New GameText
            txt1Up.InitializePixelColor White
            txt1Up.GenerateText "1Up", 1, 49
        Set score1Up = New GameText
            score1Up.SwitchJustification False, True 'Right Justified
            'reset 1Up score
            score1Up.InitializePixelColor White
            score1Up.GenerateText "00", 9, 65
    
    'initialize center map text
        'Ready! Message
        Set txtReady = New GameText
            txtReady.InitializePixelColor Yellow
            txtReady.GenerateText "Ready!", 161, 105
        'Game Over Message
        Set txtGameOver = New GameText
            txtGameOver.InitializePixelColor Red
            txtGameOver.GenerateText "Game Over", 161, 93
    
    'set level map colors
        If LvlNum = 1 Then 'Level 1
            Map.ColorOutline = Red
            Map.ColorFill = Salmon
            Map.ColorTrack = TrackBlack
'FOR DEBUGGING MOVEMENTS=======================================================================================================
'            Map.ColorTrack = Green 'For Debugging Track Range
'FOR DEBUGGING MOVEMENTS=======================================================================================================
            Map.ColorGate = Pink
            Map.ColorDots = White
        End If
    
    'assign map-range array colors
        MR_Array(1, 1) = Map.ColorOutline
        MR_Array(2, 1) = Map.ColorFill
        MR_Array(3, 1) = Map.ColorTrack
        MR_Array(4, 1) = Map.ColorGate
        MR_Array(5, 1) = Map.ColorDots
        MR_Array(6, 1) = White 'Always White as this is for the 2 header text objects
    
    'initialize map
        Map.Map_Initialize
    
    'initialize possible tiles in tile list
        Map.Tiles.Initialize_Tiles
    
    'level map seed
        If LvlNum = 1 Then 'Level 1
             strLevelSeed = "o1o o6o o6o o6o o6o o6o o6o v8o v7o o6o o6o o6o o6o o6o o6o o6o o6o o6o o6o v8o v7o o6o o6o o6o o6o o6o o6o o2o " & _
                            "o5o t1p t6p t6p t6p t6p t2p i5o i7o t1p t6p t6p t6p t6p t6p t6p t6p t6p t2p i5o i7o t1p t6p t6p t6p t6p t2p o7o " & _
                            "o5o t5P i1o i6o i6o i2o t5p i5o i7o t5p i1o i6o i6o i6o i6o i6o i6o i2o t5p i5o i7o t5p i1o i6o i6o i2o t5P o7o " & _
                            "o5o t5p i4o i8o i8o i3o t5p i4o i3o t5p i4o i8o i8o i8o i8o i8o i8o i3o t5p i4o i3o t5p i4o i8o i8o i3o t5p o7o " & _
                            "o5o t4p t6p x4p t6p t6p x5p t6p t6p x2p t6p t6p x4p t6p t6p x4p t6p t6p x2p t6p t6p x5p t6p t6p x4p t6p t3p o7o " & _
                            "o4o o8o i2o t5p i1o i2o t5p i1o i6o i6o i6o i2o t5p i1o i2o t5p i1o i6o i6o i6o i2o t5p i1o i2o t5p i1o o8o o3o " & _
                            "f0o f0o o5o t5p i5o i7o t5p i5o f1o f1o f1o i7o t5p i5o i7o t5p i5o f1o f1o f1o i7o t5p i5o i7o t5p o7o f0o f0o " & _
                            "o6o o6o i3o t5p i5o i7o t5p i4o i8o i8o i8o i3o t5p i5o i7o t5p i4o i8o i8o i8o i3o t5p i5o i7o t5p i4o o6o o6o " & _
                            "t6o t6o t6o x1p i5o i7o t4p t6p t6p x4p t6p t6p t3p i5o i7o t4p t6p t6p x4p t6p t6p t3p i5o i7o x3p t6o t6o t6o " & _
                            "o8o o8o i2o t5p i5o v2o i6o i6o i2o t5o i1o i6o i6o v1o v2o i6o i6o i2o t5o i1o i6o i6o v1o i7o t5p i1o o8o o8o " & _
                            "f0o f0o o5o t5p i4o i8o i8o i8o i3o t5o i4o i8o i8o i8o i8o i8o i8o i3o t5o i4o i8o i8o i8o i3o t5p o7o f0o f0o " & _
                            "f0o f0o o5o x3p t6o t6o t6o t6o t6o x5o t6o t6o t6o t6o t6o t6o t6o t6o x5o t6o t6o t6o t6o t6o x1p o7o f0o f0o " & _
                            "f0o f0o o5o t5p i1o i6o i6o i6o i2o t5o c1o c6o c6o g1o g2o c6o c6o c2o t5o i1o i6o i6o i6o i2o t5p o7o f0o f0o " & _
                            "f0o f0o o5o t5p i5o v3o i8o i8o i3o t5o c5o f0o f0o f0o f0o f0o f0o c7o t5o i4o i8o i8o v4o i7o t5p o7o f0o f0o " & _
                            "f0o f0o o5o t5p i5o i7o t1o t6o t6o x1o c5o f0o f0o f0o f0o f0o f0o c7o x3o t6o t6o t2o i5o i7o t5p o7o f0o f0o " & _
                            "f0o f0o o5o t5p i5o i7o t5o i1o i2o t5o c5o f0o f0o f0o f0o f0o f0o c7o t5o i1o i2o t5o i5o i7o t5p o7o f0o f0o "
            strLevelSeed = strLevelSeed & _
                            "o6o o6o i3o t5p i4o i3o t5o i5o i7o t5o c4o c8o c8o c8o c8o c8o c8o c3o t5o i5o i7o t5o i4o i3o t5p i4o o6o o6o " & _
                            "t6o t6o t6o x5p t6o t6o t3o i5o i7o t4o t6o t6o x4o t6o t6o x4o t6o t6o t3o i5o i7o t4o t6o t6o x5p t6o t6o t6o " & _
                            "o8o o8o i2o t5p i1o i6o i6o v1o v2o i6o i6o i2o t5o i1o i2o t5o i1o i6o i6o v1o v2o i6o i6o i2o t5p i1o o8o o8o " & _
                            "f0o f0o o5o t5p i4o i8o i8o i8o i8o i8o i8o i3o t5o i5o i7o t5o i4o i8o i8o i8o i8o i8o i8o i3o t5p o7o f0o f0o " & _
                            "f0o f0o o5o x3p t6p t6p t6p t6p t6p x4p t6o t6o t3o i5o i7o t4o t6o t6o x4p t6p t6p t6p t6p t6p x1p o7o f0o f0o " & _
                            "f0o f0o o5o t5p i1o i6o i6o i6o i2o t5p i1o i6o i6o v1o v2o i6o i6o i2o t5p i1o i6o i6o i6o i2o t5p o7o f0o f0o " & _
                            "o1o o6o i3o t5p i4o i8o i8o i8o i3o t5p i4o i8o i8o i8o i8o i8o i8o i3o t5p i4o i8o i8o i8o i3o t5p i4o o6o o2o " & _
                            "o5o t1p t6p x2p t6p t6p x4p t6p t6p x2p t6p t6p x4p t6o t6o x4p t6p t6p x2p t6p t6p x4p t6p t6p x2p t6p t2p o7o " & _
                            "o5o t5p i1o i6o i6o i2o t5p i1o i6o i6o i6o i2o t5p i1o i2o t5p i1o i6o i6o i6o i2o t5p i1o i6o i6o i2o t5p o7o " & _
                            "o5o t5p i5o f1o f1o i7o t5p i5o v3o i8o i8o i3o t5p i5o i7o t5p i4o i8o i8o v4o i7o t5p i5o f1o f1o i7o t5p o7o " & _
                            "o5o t5p i5o f1o f1o i7o t5p i5o i7o t1p t6p t6p t3p i5o i7o t4p t6p t6p t2p i5o i7o t5p i5o f1o f1o i7o t5p o7o " & _
                            "o5o t5P i5o f1o f1o i7o t5p i5o i7o t5p i1o i6o i6o v1o v2o i6o i6o i2o t5p i5o i7o t5p i5o f1o f1o i7o t5P o7o " & _
                            "o5o t5p i4o i8o i8o i3o t5p i4o i3o t5p i4o i8o i8o i8o i8o i8o i8o i3o t5p i4o i3o t5p i4o i8o i8o i3o t5p o7o " & _
                            "o5o t4p t6p t6p t6p t6p x2p t6p t6p x2p t6p t6p t6p t6p t6p t6p t6p t6p x2p t6p t6p x2p t6p t6p t6p t6p t3p o7o " & _
                            "o4o o8o o8o o8o o8o o8o o8o o8o o8o o8o o8o o8o o8o o8o o8o o8o o8o o8o o8o o8o o8o o8o o8o o8o o8o o8o o8o o3o"
        End If

    'load seed to map object
        Map.Seed_Load (strLevelSeed)
    
    'load level map to map object
        Map.Map_Load
    
    'draw map to drawing arrays
        Map.Map_Draw
    
    'draw level header
        txt1Up.Map_DrawText 'draw ONCE to map arrays
        txtHighScore.Map_DrawText 'draw ONCE to map arrays
        score1Up.DrawText 'draw to color arrays
        scoreHighScore.DrawText 'draw to color arrays
    
    'draw food footer
        For i = 1 To FoodTypesNumber
            FoodTypesArr(i).UpdatePosition (True)
            FoodTypesArr(i).Draw
        Next i
    
    'Grab remaining CLR_Array ranges from loading to CR_Array ranges
        For i = 1 To 9
            'dump remainder into cr_array
            Set CR_Array((i + 1), 2) = Union(CR_Array((i + 1), 2), CLR_Array(i, 2))
            'clear clr_array
            Set CLR_Array(i, 2) = Cells(2, (i + 7))
            CLR_Array(i, 1) = 0
        Next i
    
    'Grab remaining MLR_Array ranges from loading to MR_Array and CR_Array ranges
        'MR_Array ranges
            For i = 1 To UBound(MR_Array)
                Set MR_Array(i, 2) = Union(MR_Array(i, 2), MLR_Array(i, 2))
            Next i
        'CR_Array range
            Set CR_Array(1, 2) = Union(CR_Array(1, 2), MLR_Array(7, 2))
        'Clear MLR_Array rangescv and counters
            For i = 1 To UBound(MLR_Array)
                MLR_Array(i, 1) = 0
                Set MLR_Array(i, 2) = Cells(2, i)
            Next i
    
    'draw level map
        Application.ScreenUpdating = False 'turn off screen updates
        'Color Ranges
        For i = 1 To UBound(CR_Array)
            If CR_Array(i, 2).Count <> 1 Then
                CR_Array(i, 2).Interior.Color = CR_Array(i, 1) 'color ranges (track range is also a part of black range)
            End If
        Next i
        'Map Ranges
        For i = 1 To UBound(MR_Array)
            If MR_Array(i, 2).Count <> 1 Then
                MR_Array(i, 2).Interior.Color = MR_Array(i, 1)
            End If
        Next i
        Application.ScreenUpdating = True 'turn on screen updates
    
    'add (track range - pacdots range) to black range
        Set CR_Array(1, 2) = Union(CR_Array(1, 2), MR_Array(3, 2))
            'will change track color to regular black, but this is fine since
            'the track's range is really what matters in the frame updater.
            'Also, the track range can be accessed as seperate from the black range
            'which allows the track to still function as intended.
    
    'add pacdots range to track range
        Set MR_Array(3, 2) = Union(MR_Array(3, 2), MR_Array(5, 2))
            'fixes holes in the track range from pacdots, not having pacdots
            'as a part of track range before, when added to the black range,
            'allows for the pacdots to remain on the map as a part of the background.
            'Conditionals will later tell each pacdot to disappear or reappear, depending
            'on which type of sprite passes overtop of it.

End Sub

Public Sub UpdateScores(ByVal AddScore As Integer)

    'update scores
        valueHighScore = valueHighScore + AddScore
        value1Up = value1Up + AddScore
    
    '1Up Achieved
    If value1Up >= 10000 Then
        'reset value
        value1Up = value1Up - 10000
        'Initialize new extra life (max 7 lives)
        If MsPLives < 7 Then
            'play Extra Life audio
            Call PlayAudio(arrExtraLife, False)
            'give MsP a life
            MsPLives = MsPLives + 1
            ReDim Preserve MsPLivesArr(1 To MsPLives) As Sprite_MsPacman
            'initialize new life
            Set MsPLivesArr(MsPLives) = New Sprite_MsPacman
            With MsPLivesArr(MsPLives)
                .Xcol = 21 + (((MsPLives - 1) - 1) * 16) '(MspLives-1) to ignore 1st entry
                .Yrow = 273
                .nAnim = 2 'Resting Anim
                .nDir = 3 'Facing Right
                .Vx = 0 'no velocity
                .Vy = 0 'no velocity
                .Speed = 0 'zero speed
                .SetPosition
                .InitializeSprites
                .SetAnimDir
            End With
        End If
    End If
    
    'Rewrite Text Objects
        'High Score
        scoreHighScore.GenerateText CStr(valueHighScore), scoreHighScore.Yrow, scoreHighScore.Xcol
        scoreHighScore.DrawText
        '1Up Score
        score1Up.GenerateText CStr(value1Up), score1Up.Yrow, score1Up.Xcol
        score1Up.DrawText

End Sub

Public Sub AmbientAudioResetTimers_On_Off()

    'Switch Reset Timers Booleans
    arrGhostAmbient(1, 4) = Not arrGhostAmbient(1, 4)
    arrGhostAmbient(2, 4) = Not arrGhostAmbient(2, 4)
    arrFoodAmbient(1, 4) = Not arrFoodAmbient(1, 4)
    arrFoodAmbient(2, 4) = Not arrFoodAmbient(2, 4)
    
    'Reset Timers
    If arrGhostAmbient(1, 4) And arrGhostAmbient(2, 4) And arrFoodAmbient(1, 4) And arrFoodAmbient(2, 4) Then
        arrGhostAmbient(1, 3) = 2000
        arrGhostAmbient(2, 3) = 0
        arrFoodAmbient(1, 3) = 2000
        arrFoodAmbient(2, 3) = 0
    End If

End Sub

Public Sub AmbientAudioController()

    'Ghost Ambient Loop
        'Play Side A
            If arrGhostAmbient(1, 3) = 2000 And arrGhostAmbient(2, 3) = 0 Then
                Call PlayAudio(arrGhostAmbient, True, "A")
                arrGhostAmbient(1, 3) = arrGhostAmbient(1, 3) - 1
        'Decrement A
            ElseIf arrGhostAmbient(1, 3) > 0 Then
                arrGhostAmbient(1, 3) = arrGhostAmbient(1, 3) - 1
        'Flip Counters to B
            ElseIf arrGhostAmbient(1, 2) And arrGhostAmbient(1, 3) = 0 And arrGhostAmbient(2, 3) = 0 Then
                If arrGhostAmbient(2, 4) Then: arrGhostAmbient(2, 3) = 2000
        'Play Side B
            ElseIf arrGhostAmbient(1, 3) = 0 And arrGhostAmbient(2, 3) = 2000 Then
                Call PlayAudio(arrGhostAmbient, True, "B")
                arrGhostAmbient(2, 3) = arrGhostAmbient(2, 3) - 1
        'Decrement B
            ElseIf arrGhostAmbient(2, 3) > 0 Then
                arrGhostAmbient(2, 3) = arrGhostAmbient(2, 3) - 1
        'Flip Counters to A
            ElseIf arrGhostAmbient(2, 2) And arrGhostAmbient(1, 3) = 0 And arrGhostAmbient(2, 3) = 0 Then
                If arrGhostAmbient(1, 4) Then: arrGhostAmbient(1, 3) = 2000
            End If
    
    
    'Food Ambient Loop
        If Food.AmbientAudio Then
        'Play Side A
            If arrFoodAmbient(1, 3) = 2000 And arrFoodAmbient(2, 3) = 0 Then
                Call PlayAudio(arrFoodAmbient, True, "A")
                arrFoodAmbient(1, 3) = arrFoodAmbient(1, 3) - 1
        'Decrement A
            ElseIf arrFoodAmbient(1, 3) > 0 Then
                arrFoodAmbient(1, 3) = arrFoodAmbient(1, 3) - 1
        'Flip Counters to B
            ElseIf arrFoodAmbient(1, 2) And arrFoodAmbient(1, 3) = 0 And arrFoodAmbient(2, 3) = 0 Then
                If arrFoodAmbient(2, 4) Then: arrFoodAmbient(2, 3) = 2000
        'Play Side B
            ElseIf arrFoodAmbient(1, 3) = 0 And arrFoodAmbient(2, 3) = 2000 Then
                Call PlayAudio(arrFoodAmbient, True, "B")
                arrFoodAmbient(2, 3) = arrFoodAmbient(2, 3) - 1
        'Decrement B
            ElseIf arrFoodAmbient(2, 3) > 0 Then
                arrFoodAmbient(2, 3) = arrFoodAmbient(2, 3) - 1
        'Flip Counters to A
            ElseIf arrFoodAmbient(2, 2) And arrFoodAmbient(1, 3) = 0 And arrFoodAmbient(2, 3) = 0 Then
                If arrFoodAmbient(1, 4) Then: arrFoodAmbient(1, 3) = 2000
            End If
        'Stop Playing food ambient audio
        Else
            Call StopAudio(arrFoodAmbient, True, "A")
            Call StopAudio(arrFoodAmbient, True, "B")
        End If


End Sub

Public Sub LoadWaveIntoMemory(sWaveFile$)
'Loads a wav file to memory into its corresponding byte array

    Dim FLength&, hOrgFile&, Ret&
    
    'get file length in bytes
    FLength = FileLen(sWaveFile)
    
    'resize byte array
    ReDim bytearrChomp(1 To FLength)
    
    'create file
    hOrgFile = CreateFile(sWaveFile, GENERIC_READ, ByVal 0&, ByVal 0&, OPEN_EXISTING, 0, 0)
    
    'read file
    ReadFile hOrgFile, bytearrChomp(1), FLength, Ret, ByVal 0&
    
    'Error
    If Ret <> FLength Then
        MsgBox "Error while reading file"
    End If
    
    'close file
    CloseHandle hOrgFile

End Sub

Public Sub PlayByteArray_Chomp()
'Plays byte array wave file from memory
    
    'play wav file from memory
    PlaySound bytearrChomp(1), ByVal 0&, SND_MEMORY Or SND_ASYNC

End Sub

Public Sub PlayAudio(ByRef AudioArray As Variant, ByVal AmbientArray As Boolean, Optional ByVal AorB As String)
'Play Most Audio Files via audio arrays


    'non ambient arrays
    If Not AmbientArray Then
        'change playing boolean
        AudioArray(2) = True
        'play audio
        Call mciSendString("play " & AudioArray(1), 0&, 0, 0)
    
    
    'ambient arrays
    ElseIf AmbientArray Then
        'A
        If AorB = "A" Then
            'change playing booleans
            AudioArray(1, 2) = True
            If AudioArray(2, 2) Then: Call StopAudio(AudioArray, True, "B")
            AudioArray(2, 2) = False
            'play audio
            Call mciSendString("play " & AudioArray(1, 1), 0&, 0, 0)
        'B
        ElseIf AorB = "B" Then
            'change playing booleans
            If AudioArray(1, 2) Then: Call StopAudio(AudioArray, True, "A")
            AudioArray(1, 2) = False
            AudioArray(2, 2) = True
            'play audio
            Call mciSendString("play " & AudioArray(2, 1), 0&, 0, 0)
        End If
    
    
    End If

End Sub

Public Sub StopAudio(ByRef AudioArray As Variant, ByVal AmbientArray As Boolean, Optional ByVal AorB As String)
'Stop Most Audio Files via audio arrays


    'non ambient arrays
    If Not AmbientArray Then
        'change playing boolean
        AudioArray(2) = False
        'stop audio
        Call mciSendString("stop " & AudioArray(1), 0&, 0, 0)
    
    
    'ambient arrays
    ElseIf AmbientArray Then
        'A
        If AorB = "A" Then
            'change playing boolean
            AudioArray(1, 2) = False
            'reset repeat timers
            If AudioArray(1, 4) And AudioArray(2, 4) Then
                AudioArray(1, 3) = 2000
                AudioArray(2, 3) = 0
            Else
                AudioArray(1, 3) = 0
                AudioArray(2, 3) = 0
            End If
            'stop audio
            Call mciSendString("stop " & AudioArray(1, 1), 0&, 0, 0)
        'B
        ElseIf AorB = "B" Then
            'change playing boolean
            AudioArray(2, 2) = False
            'reset repeat timers
            If AudioArray(1, 4) And AudioArray(2, 4) Then
                AudioArray(1, 3) = 2000
                AudioArray(2, 3) = 0
            Else
                AudioArray(1, 3) = 0
                AudioArray(2, 3) = 0
            End If
            'stop audio
            Call mciSendString("stop " & AudioArray(2, 1), 0&, 0, 0)
        End If
    
    
    End If

End Sub

Public Sub EndAudio(ByRef AudioArray As Variant, ByVal AmbientArray As Boolean, Optional ByVal AorB As String)
'End Most Audio Files via audio arrays for end of game


    'non ambient arrays
    If Not AmbientArray Then
        'change playing boolean
        AudioArray(2) = False
        'stop audio
        Call mciSendString("stop " & AudioArray(1), 0&, 0, 0)
    
    
    'ambient arrays
    ElseIf AmbientArray Then
        'A
        If AorB = "A" Then
            'change playing boolean
            AudioArray(1, 2) = False
            'reset repeat timers
            AudioArray(1, 3) = -1
            AudioArray(2, 3) = -1
            'stop audio
            Call mciSendString("stop " & AudioArray(1, 1), 0&, 0, 0)
        'B
        ElseIf AorB = "B" Then
            'change playing boolean
            AudioArray(2, 2) = False
            'reset repeat timers
            AudioArray(1, 3) = -1
            AudioArray(2, 3) = -1
            'play audio
            Call mciSendString("stop " & AudioArray(2, 1), 0&, 0, 0)
        End If
    
    
    End If

End Sub
