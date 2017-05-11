Attribute VB_Name = "GameOfLife"
'=====================================Declarations=======================================

'using kernel32 API
Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

'===================================Public Variables=====================================

Public Wrkb As Workbook 'Workbook in which the game of life occurs
Public Wrks As Worksheet 'Worksheet on which the game of life occurs
Public ControlBoard As GoLControls 'Userform that controls game of life
Public PixelEditing As Boolean 'boolean to turn pixel editing on or off
Public GameInProgress As Boolean 'boolean to tell excel the game is currently running
Public GameRange As Range 'range holding Game Grid Area
Public GameGrid() As Variant 'array holding all pixel objects
Public Red As Long 'color of game grid border
Public White As Long 'color of dead pixels
Public Black As Long 'color of alive pixels

'======================================Subroutines========================================

Public Sub TheGameOfLife()

    'initialize worksheet variables
    Set Wrkb = Workbooks("Game of Life")
    Set Wrks = Wrkb.Worksheets("Game Of Life")
    
    'initialize userform
    Set ControlBoard = New GoLControls
    
    'move selection out of the way
    Wrks.Range("A1").Select
    
    'show userform
    ControlBoard.Show vbModeless
    DoEvents

End Sub

Public Sub PixelEditor_On_Off()

    'switch boolean
    If PixelEditing Then
        PixelEditing = False
    Else
        PixelEditing = True
    End If

End Sub

Public Sub Initialize()

    Dim c As Range 'generic range iteration object
    
    
    'initialize variables
    GameInProgress = True
    Red = RGB(255, 0, 0)
    White = RGB(255, 255, 255)
    Black = RGB(0, 0, 0)
    
    
    'initialize game range
'    Set GameRange = Wrks.Range(Cells(2, 2), Cells(51, 131))
'    ReDim GameGrid(1 To 130, 1 To 50)
    Set GameRange = Wrks.Range(Cells(2, 2), Cells(51, 51))
    ReDim GameGrid(1 To 50, 1 To 50)
    
        
    'Read Grid
    Call ReadGrid


End Sub

Public Sub Start()

    'move selection out of the way
    Wrks.Range("A1").Select
    
    'call initializer
    Call Initialize
    
    'game in progress loop
    Application.Calculation = xlCalculationManual
    Do Until Not GameInProgress
        DoEvents
        'call counter
        Call Counter
        'sleep for 0.001 seconds
        Sleep 1
    Loop
    Application.Calculation = xlCalculationAutomatic

End Sub

Public Sub Pause()

    'only pause if game is running
    If GameInProgress Then
        'change boolean
        GameInProgress = False
    End If

End Sub

Public Sub Clear()

    'call initializer
    Call Initialize
    
    'only pause if game in progress
    If GameInProgress Then: Call Pause
        
    'clear board
    GameRange.Interior.Color = xlNone

End Sub

Public Sub Counter()

    Application.ScreenUpdating = False
    
    'Calculate New Grid
    Call CalculateNextGeneration
    
    'Write Grid
    Call WriteGrid
    
    Application.ScreenUpdating = True

End Sub

Public Sub CalculateNextGeneration()

    Dim tPixel As Pixel 'temp pixel object for grabbing from GameGrid Array
    
    'iterate game grid array
    For Each pxl In GameGrid 'pxl is a undefined variable that gets its data type as well as its value from each entry in GameGrid() by the for loop
        Set tPixel = pxl
        tPixel.ReadNeighbors
        tPixel.Evolve
    Next pxl
    
End Sub

Public Sub WriteGrid()

    Dim c As Range 'range iteration object

    'iterate game range
    For Each c In GameRange
        'living pixel
        If GameGrid((c.Column - 1), (c.Row - 1)).Alive Then
            c.Interior.Color = Black
        'dead pixel
        ElseIf Not GameGrid((c.Column - 1), (c.Row - 1)).Alive Then
            c.Interior.Color = xlNone
        End If
    Next c

End Sub

Public Sub ReadGrid()
    
    Dim tPixel As Pixel 'temp pixel object for placing into GameGrid Array

    'create and place pixel objects into GameGrid array
    For Each c In GameRange

        'create pixel
        Set tPixel = New Pixel
        'pixel location
        tPixel.Xcol = c.Column
        tPixel.Yrow = c.Row
        'pixel alive or dead
        If c.Interior.Color = Black Then
            tPixel.Alive = True
        ElseIf c.Interior.Color = White Then
            tPixel.Alive = False
        End If

        'place temp pixel into array
        Set GameGrid((tPixel.Xcol - 1), (tPixel.Yrow - 1)) = tPixel

    Next c

End Sub
