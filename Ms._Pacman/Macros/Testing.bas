Attribute VB_Name = "Testing"
Declare PtrSafe Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer
Private VK_UP As Long
Private VK_LEFT As Long
Private VK_RIGHT As Long
Private VK_DOWN As Long
Private VK_W As Long
Private VK_A As Long
Private VK_D As Long
Private VK_S As Long

Public Sub Demo()
    Dim lngResult As Long
'    lngResult = mciSendString("play " & "C:\Users\xrgb231\Desktop\For_Portfolio\Video_Game_Programming\MsP\Audio\wav16bit\MsP_FoodAmbient.wav", 0&, 0, 0)
    lngResult = mciSendString("play " & "C:\Users\xrgb231\Desktop\For_Portfolio\Video_Game_Programming\MsP\Audio\wav16bit\MsP_GhostAmbientB.wav", 0&, 0, 0)
'    lngResult = mciSendString("play " & "C:\Users\xrgb231\Desktop\For_Portfolio\Video_Game_Programming\MsP\Audio\wav16bit\MsP_PowerDotAmbient.wav", 0&, 0, 0)
End Sub

Public Sub Demo2()

    Dim arrTest(1 To 2, 1 To 3) As Variant
    Dim i As Long
    
    arrTest(1, 1) = "C:\Users\xrgb231\Desktop\For_Portfolio\Video_Game_Programming\MsP\Audio\wav16bit\MsP_GhostAmbientA.wav"
    arrTest(1, 2) = False
    arrTest(1, 3) = 2000
    arrTest(2, 1) = "C:\Users\xrgb231\Desktop\For_Portfolio\Video_Game_Programming\MsP\Audio\wav16bit\MsP_GhostAmbientB.wav"
    arrTest(2, 2) = False
    arrTest(2, 3) = 0
    
    For i = 1 To 10000
    
        If arrTest(1, 3) = 2000 And arrTest(2, 3) = 0 Then
            Call PlayAudio(arrTest, True, "A")
            arrTest(1, 3) = arrTest(1, 3) - 1
        ElseIf arrTest(1, 3) > 0 Then
            arrTest(1, 3) = arrTest(1, 3) - 1
        ElseIf arrTest(1, 2) And arrTest(1, 3) = 0 And arrTest(2, 3) = 0 Then
            arrTest(2, 3) = 2000
        
        ElseIf arrTest(1, 3) = 0 And arrTest(2, 3) = 2000 Then
            Call PlayAudio(arrTest, True, "B")
            arrTest(2, 3) = arrTest(2, 3) - 1
        ElseIf arrTest(2, 3) > 0 Then
            arrTest(2, 3) = arrTest(2, 3) - 1
        ElseIf arrTest(2, 2) And arrTest(1, 3) = 0 And arrTest(2, 3) = 0 Then
            arrTest(1, 3) = 2000
        End If
    
    Next i

End Sub

Public Function PlayMultimedia(AliasName As String, from_where As String, to_where As String) As String
'Calling PlayMultimedia will play the multimedia file.

'Parameters:
    'AliasName
        '[in]Specifies alias name of the file you want to play
    'from_where
        '[in]Specifies what frame to start playing the multimedia file at
    'to_where
        '[in]Specifies what frame to stop playing the multimedia file at
    
    Dim cmdToDo As String 'string to give to mciSendString function
    Dim dwReturn As Long 'long to receive from mciSendString function
    Dim Ret As String '?????
    
    If from_where = vbNullString Then: from_where = 0
    If to_where = vbNullString Then: to_where = GetTotalFrames(AliasName)
    
    'Important for auto repeat
    If AliasName = glo_AliasName Then
        glo_from = from_where
        glo_to = to_where
    End If
    
    cmdToDo = "play " & AliasName '& " from " & from_where & " to " & to_where
    
    dwReturn = mciSendString(cmdToDo, 0&, 0, 0) 'playing multimedia file
    
    If Not dwReturn = 0 Then 'play error
        mciGetErrorString dwReturn, Ret, 128 'get error message
        PlayMultimedia = Ret 'return the error
        Exit Function
    End If
    
    'Success
    PlayMultimedia = "Success"

End Function

Public Function GetTotalFrames(AliasName As String) As Long
'Get the total number of frames for the multimedia file

'Parameters:
    'AliasName
        '[in]Specifies alias name of multimedia file

    Dim dwReturn As Long
    Dim Total As String * 128
    
    If Not dwReturn = 0 Then 'error
        GetTotalFrames = -1
        Exit Function
    End If
    
    'Success
    GetTotalFrames = Val(Total)

End Function

Public Sub test_TextObjects()

    Dim tText As GameText
    Dim ttext2 As GameText
    
    Set tText = New GameText
    Set ttext2 = New GameText
    
    tText.GenerateText "hello world!", RGB(255, 255, 255), 100, 100
    ttext2.GenerateText "Pls woRK! 0123456789", RGB(255, 0, 0), 132, 100
    
    tText.DrawText
    ttext2.DrawText

End Sub

Public Sub test_GetKeyState()
'More complicated than this BUT: keys are negative when held down (-127 or -128) and positive when up (0 or 1)

    Dim intUp As Integer
    Dim intLeft As Integer
    Dim intRight As Integer
    Dim intDown As Integer
    Dim intW As Integer
    Dim intA As Integer
    Dim intD As Integer
    Dim intS As Integer
    
    VK_UP = 38
    VK_LEFT = 37
    VK_RIGHT = 39
    VK_DOWN = 40
    VK_W = 87
    VK_A = 65
    VK_D = 68
    VK_S = 83
    
    intUp = GetKeyState(VK_UP)
    intLeft = GetKeyState(VK_LEFT)
    intRight = GetKeyState(VK_RIGHT)
    intDown = GetKeyState(VK_DOWN)
    intW = GetKeyState(VK_W)
    intA = GetKeyState(VK_A)
    intD = GetKeyState(VK_D)
    intS = GetKeyState(VK_S)
    
    MsgBox "Up:  " & intUp & vbNewLine & _
            "Left:  " & intLeft & vbNewLine & _
            "Right:  " & intRight & vbNewLine & _
            "Down:  " & intDown & vbNewLine & _
            "W:  " & intW & vbNewLine & _
            "A:  " & intA & vbNewLine & _
            "D:  " & intD & vbNewLine & _
            "S:  " & intS

End Sub

Public Sub test_ScrollAnimation()

    Dim i As Integer
    
    GameInProgress = True
    i = 0
    
    'game in progress loop
    Do Until Not GameInProgress
        DoEvents
        
        Application.Goto Cells(((i * 300) + 1), 1), True
        i = i + 1
        
        If i = 8 Then i = 0
        
    Loop

End Sub

Public Sub Fix_TrackColorLvl1()

    Dim c As Range 'iterator
    
    Application.ScreenUpdating = False
    
    For Each c In Range(Cells(1, 1), Cells(300, 280))
        If c.Interior.Color = RGB(0, 255, 0) Then: c.Interior.Color = RGB(15, 36, 62)
        If c.Interior.Color = RGB(254, 254, 254) Then: c.Interior.Color = RGB(204, 192, 218) '1 less in each for places where the track transects the dots
    Next c
    
    Application.ScreenUpdating = True

End Sub

Public Sub test_MapClass()

    Dim tMap As Sprite_Map 'map object
    Dim strLevelSeed As String 'level string to be given to map's Seed_Load method
    Dim i As Integer
    
    'intialize map object
    Set tMap = New Sprite_Map
    
    'set level map colors
    tMap.ColorOutline = Red
    tMap.ColorFill = Salmon
    tMap.ColorGate = Pink
    
    'set level map dimensions
    tMap.WidthCubes = 28
    tMap.HeightCubes = 31
    
    'initialize map
    tMap.Map_Initialize
    
    'initialize possible tiles in tile list
    tMap.Tiles.Initialize_Tiles
    
    'level map seed
     strLevelSeed = "o1 o6 o6 o6 o6 o6 o6 v8 v7 o6 o6 o6 o6 o6 o6 o6 o6 o6 o6 v8 v7 o6 o6 o6 o6 o6 o6 o2 " & _
                    "o5 b1 b1 b1 b1 b1 b1 i5 i7 b1 b1 b1 b1 b1 b1 b1 b1 b1 b1 i5 i7 b1 b1 b1 b1 b1 b1 o7 " & _
                    "o5 b1 i1 i6 i6 i2 b1 i5 i7 b1 i1 i6 i6 i6 i6 i6 i6 i2 b1 i5 i7 b1 i1 i6 i6 i2 b1 o7 " & _
                    "o5 b1 i4 i8 i8 i3 b1 i4 i3 b1 i4 i8 i8 i8 i8 i8 i8 i3 b1 i4 i3 b1 i4 i8 i8 i3 b1 o7 " & _
                    "o5 b1 b1 b1 b1 b1 b1 b1 b1 b1 b1 b1 b1 b1 b1 b1 b1 b1 b1 b1 b1 b1 b1 b1 b1 b1 b1 o7 " & _
                    "o4 o8 i2 b1 i1 i2 b1 i1 i6 i6 i6 i2 b1 i1 i2 b1 i1 i6 i6 i6 i2 b1 i1 i2 b1 i1 o8 o3 " & _
                    "b1 b1 o5 b1 i5 i7 b1 i5 f1 f1 f1 i7 b1 i5 i7 b1 i5 f1 f1 f1 i7 b1 i5 i7 b1 o7 b1 b1 " & _
                    "o6 o6 i3 b1 i5 i7 b1 i4 i8 i8 i8 i3 b1 i5 i7 b1 i4 i8 i8 i8 i3 b1 i5 i7 b1 i4 o6 o6 " & _
                    "b1 b1 b1 b1 i5 i7 b1 b1 b1 b1 b1 b1 b1 i5 i7 b1 b1 b1 b1 b1 b1 b1 i5 i7 b1 b1 b1 b1 " & _
                    "o8 o8 i2 b1 i5 v2 i6 i6 i2 b1 i1 i6 i6 v1 v2 i6 i6 i2 b1 i1 i6 i6 v1 i7 b1 i1 o8 o8 " & _
                    "b1 b1 o5 b1 i4 i8 i8 i8 i3 b1 i4 i8 i8 i8 i8 i8 i8 i3 b1 i4 i8 i8 i8 i3 b1 o7 b1 b1 " & _
                    "b1 b1 o5 b1 b1 b1 b1 b1 b1 b1 b1 b1 b1 b1 b1 b1 b1 b1 b1 b1 b1 b1 b1 b1 b1 o7 b1 b1 " & _
                    "b1 b1 o5 b1 i1 i6 i6 i6 i2 b1 c1 c6 c6 g1 g2 c6 c6 c2 b1 i1 i6 i6 i6 i2 b1 o7 b1 b1 " & _
                    "b1 b1 o5 b1 i5 v3 i8 i8 i3 b1 c5 b1 b1 b1 b1 b1 b1 c7 b1 i4 i8 i8 v4 i7 b1 o7 b1 b1 " & _
                    "b1 b1 o5 b1 i5 i7 b1 b1 b1 b1 c5 b1 b1 b1 b1 b1 b1 c7 b1 b1 b1 b1 i5 i7 b1 o7 b1 b1 " & _
                    "b1 b1 o5 b1 i5 i7 b1 i1 i2 b1 c5 b1 b1 b1 b1 b1 b1 c7 b1 i1 i2 b1 i5 i7 b1 o7 b1 b1 "
    strLevelSeed = strLevelSeed & _
                    "o6 o6 i3 b1 i4 i3 b1 i5 i7 b1 c4 c8 c8 c8 c8 c8 c8 c3 b1 i5 i7 b1 i4 i3 b1 i4 o6 o6 " & _
                    "b1 b1 b1 b1 b1 b1 b1 i5 i7 b1 b1 b1 b1 b1 b1 b1 b1 b1 b1 i5 i7 b1 b1 b1 b1 b1 b1 b1 " & _
                    "o8 o8 i2 b1 i1 i6 i6 v1 v2 i6 i6 i2 b1 i1 i2 b1 i1 i6 i6 v1 v2 i6 i6 i2 b1 i1 o8 o8 " & _
                    "b1 b1 o5 b1 i4 i8 i8 i8 i8 i8 i8 i3 b1 i5 i7 b1 i4 i8 i8 i8 i8 i8 i8 i3 b1 o7 b1 b1 " & _
                    "b1 b1 o5 b1 b1 b1 b1 b1 b1 b1 b1 b1 b1 i5 i7 b1 b1 b1 b1 b1 b1 b1 b1 b1 b1 o7 b1 b1 " & _
                    "b1 b1 o5 b1 i1 i6 i6 i6 i2 b1 i1 i6 i6 v1 v2 i6 i6 i2 b1 i1 i6 i6 i6 i2 b1 o7 b1 b1 " & _
                    "o1 o6 i3 b1 i4 i8 i8 i8 i3 b1 i4 i8 i8 i8 i8 i8 i8 i3 b1 i4 i8 i8 i8 i3 b1 i4 o6 o2 " & _
                    "o5 b1 b1 b1 b1 b1 b1 b1 b1 b1 b1 b1 b1 b1 b1 b1 b1 b1 b1 b1 b1 b1 b1 b1 b1 b1 b1 o7 " & _
                    "o5 b1 i1 i6 i6 i2 b1 i1 i6 i6 i6 i2 b1 i1 i2 b1 i1 i6 i6 i6 i2 b1 i1 i6 i6 i2 b1 o7 " & _
                    "o5 b1 i5 f1 f1 i7 b1 i5 v3 i8 i8 i3 b1 i5 i7 b1 i4 i8 i8 v4 i7 b1 i5 f1 f1 i7 b1 o7 " & _
                    "o5 b1 i5 f1 f1 i7 b1 i5 i7 b1 b1 b1 b1 i5 i7 b1 b1 b1 b1 i5 i7 b1 i5 f1 f1 i7 b1 o7 " & _
                    "o5 b1 i5 f1 f1 i7 b1 i5 i7 b1 i1 i6 i6 v1 v2 i6 i6 i2 b1 i5 i7 b1 i5 f1 f1 i7 b1 o7 " & _
                    "o5 b1 i4 i8 i8 i3 b1 i4 i3 b1 i4 i8 i8 i8 i8 i8 i8 i3 b1 i4 i3 b1 i4 i8 i8 i3 b1 o7 " & _
                    "o5 b1 b1 b1 b1 b1 b1 b1 b1 b1 b1 b1 b1 b1 b1 b1 b1 b1 b1 b1 b1 b1 b1 b1 b1 b1 b1 o7 " & _
                    "o4 o8 o8 o8 o8 o8 o8 o8 o8 o8 o8 o8 o8 o8 o8 o8 o8 o8 o8 o8 o8 o8 o8 o8 o8 o8 o8 o3"

    'load seed to map object
    tMap.Seed_Load (strLevelSeed)
    
    'load level map to map object
    tMap.Map_Load
    
    'draw map to drawer
    tMap.Map_Draw
    
    'draw level map
    For i = 1 To UBound(MR_Array)
        MR_Array(i, 2).Interior.Color = MR_Array(i, 1)
    Next i

End Sub

Public Sub test_UnionRangeColors()

    Dim tR As Range
    Dim ix As Integer
    Dim iy As Integer
    
    For ix = 1 To 224
        For iy = 1 To 288
            If (ix = 1) And (iy = 1) Then
                Set tR = Cells(iy, ix)
                Set tR2 = Cells(iy, ix)
            Else
                Set tR = Union(tR, Cells(iy, ix))
            End If
        Next iy
    Next ix
    
    tR.Interior.Color = RGB(255, 0, 0)

End Sub

Public Sub test_Canvas()

    Dim tCanvas As Canvas
    
    Call MsPacman
    
    Set tCanvas = New Canvas

End Sub

Public Sub test_pixelmovement()

    Dim i As Integer
    Dim tChart As Chart
    
    Set tChart = ActiveSheet.ChartObjects("MsP Testing").Chart
    
    For i = 1 To 224
        Application.ScreenUpdating = True
        tChart.SeriesCollection(1).XValues = i
        Application.ScreenUpdating = False
    Next i

End Sub

Public Sub test_Chart()

    Dim tShp As Shape
    Dim tChart As Chart
    Dim tSeries As Series
    Dim arrXv(1 To 8, 1 To 8) As Integer
    Dim arrYv(1 To 8, 1 To 8) As Integer
    Dim i As Integer
    Dim j As Integer
    
    For i = 1 To 8
        For j = 1 To 8
            arrXv(i, j) = j
            arrYv(i, j) = i
        Next j
    Next i
    
    'create chart (28*10 by 36*10)
    Set tShp = ActiveSheet.Shapes.AddChart(xlXYScatter, 350, 15, (28 * 10), (36 * 10))
    Set tChart = tShp.Chart
    
    'set name
    tShp.Name = "MsP Testing"
    
    'change color
    tChart.ChartArea.Format.Fill.ForeColor.RGB = RGB(0, 0, 0)
    
    'add data series
    Set tSeries = tChart.SeriesCollection.NewSeries
    tSeries.XValues = arrXv
    tSeries.Values = arrYv
'    tSeries.Points(1).Format.Fill.ForeColor.RGB = RGB(255, 0, 0)
    
    'format data series
    tChart.ChartGroups(1).VaryByCategories = True
    With tSeries
        .MarkerStyle = 1
        .MarkerSize = 2
        .Format.Line.Weight = 0
        .Format.Line.Visible = msoFalse
        .Format.Fill.Visible = msoTrue
        .Format.Fill.Solid
        .Points(1).Format.Fill.ForeColor.RGB = RGB(0, 0, 0)
        .Points(2).Format.Fill.ForeColor.RGB = RGB(255, 0, 0)
        .Points(3).Format.Fill.ForeColor.RGB = RGB(0, 255, 0)
        .Points(4).Format.Fill.ForeColor.RGB = RGB(0, 0, 255)
        .Points(5).Format.Fill.ForeColor.RGB = RGB(255, 255, 255)
    End With
    
    'plot area color
    tChart.PlotArea.Format.Fill.ForeColor.RGB = RGB(0, 0, 0)
    
    'delete legend
    tChart.Legend.Delete
    
    'size plot area
    tChart.PlotArea.Select
    With Selection
        .left = -119
        .top = -24
        .Width = 556
        .Height = 541
    End With
    
    'size y axis
    With tChart.Axes(xlValue)
        .ReversePlotOrder = True
        .MinimumScale = -0.5 'start at .5 so pixels are centered when at integer values (-.5 is for extra space around borders)
        .MaximumScale = 265.5 'w = 264, extra 1.5 is for centering(.5) and border space(1)
        .MajorUnit = 1
        .MinorUnit = 1
        .MajorGridlines.Delete
        .Select
        Selection.TickLabelPosition = xlNone
        Selection.MajorTickMark = xlNone
    End With
    
    'size x axis
    With tChart.Axes(xlCategory)
        .MinimumScale = -0.5 'start at .5 so pixels are centered when at integer values (-.5 is for extra space around borders)
        .MaximumScale = 225.5 'w = 224, extra 1.5 is for centering(.5) and border space(1)
        .MajorUnit = 1
        .MinorUnit = 1
        .MajorGridlines.Delete
        .Select
        Selection.TickLabelPosition = xlNone
        Selection.MajorTickMark = xlNone
    End With
    
    'select cell A1
    Range("A1").Select

End Sub

Sub MPTest()

    'mpm obj
    Dim tmp As Sprite_MsPacman
    
    'set color variables
    Black = RGB(0, 0, 0)
    Yellow = RGB(255, 255, 0)
    Red = RGB(255, 0, 0)
    Blue = RGB(0, 0, 255)
    
    'initialize mpm
    Set tmp = New Sprite_MsPacman
    
    'set mpm coords
    tmp.Xcol = 1
    tmp.Yrow = 1
    
    'create sprites
    tmp.InitializeSprites
    
    'change animation number
'    tMP.nAnim = 1
    tmp.nAnim = 2
'    tMP.nAnim = 3
    
    'assign sprites
    tmp.SetAnimDir
    
    'draw sprites
    Application.ScreenUpdating = False
    tmp.Draw
    Application.ScreenUpdating = True

End Sub

Sub CubeTest()

    Dim tCube As Cube
    Dim tPix As Pixel
    Dim i As Integer
    Dim b As Boolean
    Dim bEven As Boolean
    
    Set tCube = New Cube
    Set tPix = New Pixel
    
    tCube.Xcol = 1
    tCube.Yrow = 1
    tCube.SetPixels
    
    b = True
    
    For i = 1 To 64
        'flip b boolean
        If (i = 9) Or (i = 17) Or (i = 25) Or (i = 33) Or (i = 41) Or (i = 49) Or (i = 57) Then
            If b Then
                b = False
            Else
                b = True
            End If
        End If
        'is i even?
        If (i / 2) = Int(i / 2) Then
            bEven = True
        Else
            bEven = False
        End If
        'set pixel color
        Set tPix = tCube.Pixels(i)
        If b Then
            If bEven Then
                tPix.Color = RGB(0, 0, 0)
            Else
                tPix.Color = RGB(255, 0, 0)
            End If
        Else
            If Not bEven Then
                tPix.Color = RGB(0, 0, 0)
            Else
                tPix.Color = RGB(255, 0, 0)
            End If
        End If
        'rewrite pixel in collection
        tCube.Pixels.Remove (i)
        If i <> 64 Then 'add in front of position i except at i=64 where there are only 63 entries after removed entry
            tCube.Pixels.Add tPix, , i
        Else 'here we can just end since its correct position is at the end
            tCube.Pixels.Add tPix
        End If
    Next i
    
    For i = 1 To 64
        Set tPix = tCube.Pixels.Item(i)
        Cells(tPix.Yrow, tPix.Xcol).Interior.Color = tPix.Color
    Next i

End Sub

Sub DrawMsP()

    'cube object
    Dim C1 As Cube
    Dim C2 As Cube
    Dim C3 As Cube
    Dim C4 As Cube
    
    'color variables
    Dim Black As Long
    Dim Yellow As Long
    Dim Red As Long
    Dim Blue As Long
    
    'pixel object property edits variables
    Dim tCube As Cube
    Dim ic As Integer
    Dim tPix As Pixel
    
    'set color variables
    Black = RGB(0, 0, 0)
    Yellow = RGB(255, 255, 0)
    Red = RGB(255, 0, 0)
    Blue = RGB(0, 0, 255)
    
    'set worksheet variables
    Set Wrkb = Workbooks("Ms. Pacman")
    Set Wrks = Wrkb.Worksheets("Ms. Pacman")
    
    'initialize cubes
    Set C1 = New Cube
    Set C2 = New Cube
    Set C3 = New Cube
    Set C4 = New Cube
    
    'cube coordinates and cube pixels
        'cube 1
        C1.Xcol = 1
        C1.Yrow = 1
        C1.SetPixels
        'cube 2
        C2.Xcol = 9
        C2.Yrow = 1
        C2.SetPixels
        'cube 3
        C3.Xcol = 1
        C3.Yrow = 9
        C3.SetPixels
        'cube 4
        C4.Xcol = 9
        C4.Yrow = 9
        C4.SetPixels
    
    'set colors
        'cube 1
            'row 3
            C1.Pixels(3, 6).Color = Yellow
            C1.Pixels(3, 7).Color = Yellow
            C1.Pixels(3, 8).Color = Yellow
            'row 4
            C1.Pixels(4, 4).Color = Yellow
            C1.Pixels(4, 5).Color = Yellow
            C1.Pixels(4, 6).Color = Yellow
            C1.Pixels(4, 7).Color = Yellow
            C1.Pixels(4, 8).Color = Yellow
            'row 5
            C1.Pixels(5, 3).Color = Yellow
            C1.Pixels(5, 4).Color = Yellow
            C1.Pixels(5, 5).Color = Yellow
            C1.Pixels(5, 6).Color = Yellow
            C1.Pixels(5, 7).Color = Yellow
            C1.Pixels(5, 8).Color = Yellow
            'row 6
            C1.Pixels(6, 3).Color = Red
            C1.Pixels(6, 4).Color = Red
            C1.Pixels(6, 5).Color = Yellow
            C1.Pixels(6, 6).Color = Yellow
            C1.Pixels(6, 7).Color = Yellow
            C1.Pixels(6, 8).Color = Black
            'row 7
            C1.Pixels(7, 7).Color = Yellow
            C1.Pixels(7, 8).Color = Yellow
        'cube 2
            'row 2
            C2.Pixels(2, 3).Color = Red
            C2.Pixels(2, 4).Color = Red
            'row 3
            C2.Pixels(3, 1).Color = Yellow
            C2.Pixels(3, 2).Color = Red
            C2.Pixels(3, 3).Color = Red
            C2.Pixels(3, 4).Color = Red
            'row 4
            C2.Pixels(4, 1).Color = Yellow
            C2.Pixels(4, 2).Color = Red
            C2.Pixels(4, 3).Color = Red
            C2.Pixels(4, 4).Color = Blue
            C2.Pixels(4, 5).Color = Red
            'row 5
            C2.Pixels(5, 1).Color = Yellow
            C2.Pixels(5, 2).Color = Yellow
            C2.Pixels(5, 3).Color = Yellow
            C2.Pixels(5, 4).Color = Red
            C2.Pixels(5, 5).Color = Blue
            C2.Pixels(5, 6).Color = Red
            C2.Pixels(5, 7).Color = Red
            'row 6
            C2.Pixels(6, 1).Color = Black
            C2.Pixels(6, 2).Color = Yellow
            C2.Pixels(6, 3).Color = Yellow
            C2.Pixels(6, 4).Color = Yellow
            C2.Pixels(6, 5).Color = Red
            C2.Pixels(6, 6).Color = Red
            C2.Pixels(6, 7).Color = Red
            'row 7
            C2.Pixels(7, 1).Color = Blue
            C2.Pixels(7, 2).Color = Black
            C2.Pixels(7, 3).Color = Yellow
            C2.Pixels(7, 4).Color = Yellow
            C2.Pixels(7, 5).Color = Red
            C2.Pixels(7, 6).Color = Red
            'row 8
            C2.Pixels(8, 1).Color = Yellow
            C2.Pixels(8, 2).Color = Yellow
            C2.Pixels(8, 3).Color = Yellow
            C2.Pixels(8, 4).Color = Yellow
            C2.Pixels(8, 5).Color = Yellow
            C2.Pixels(8, 6).Color = Yellow
        'cube 3
            'row 3
            C3.Pixels(3, 7).Color = Yellow
            C3.Pixels(3, 8).Color = Yellow
            'row 4
            C3.Pixels(4, 3).Color = Red
            C3.Pixels(4, 4).Color = Red
            C3.Pixels(4, 5).Color = Yellow
            C3.Pixels(4, 6).Color = Yellow
            C3.Pixels(4, 7).Color = Yellow
            C3.Pixels(4, 8).Color = Yellow
            'row 5
            C3.Pixels(5, 3).Color = Yellow
            C3.Pixels(5, 4).Color = Yellow
            C3.Pixels(5, 5).Color = Yellow
            C3.Pixels(5, 6).Color = Yellow
            C3.Pixels(5, 7).Color = Yellow
            C3.Pixels(5, 8).Color = Yellow
            'row 6
            C3.Pixels(6, 4).Color = Yellow
            C3.Pixels(6, 5).Color = Yellow
            C3.Pixels(6, 6).Color = Yellow
            C3.Pixels(6, 7).Color = Yellow
            C3.Pixels(6, 8).Color = Yellow
            'row 7
            C3.Pixels(7, 6).Color = Yellow
            C3.Pixels(7, 7).Color = Yellow
            C3.Pixels(7, 8).Color = Yellow
        'cube 4
            'row 1
            C4.Pixels(1, 3).Color = Yellow
            C4.Pixels(1, 4).Color = Yellow
            C4.Pixels(1, 5).Color = Yellow
            C4.Pixels(1, 6).Color = Yellow
            'row 2
            C4.Pixels(2, 1).Color = Yellow
            C4.Pixels(2, 2).Color = Yellow
            C4.Pixels(2, 3).Color = Yellow
            C4.Pixels(2, 4).Color = Yellow
            C4.Pixels(2, 5).Color = Yellow
            C4.Pixels(2, 6).Color = Yellow
            'row 3
            C4.Pixels(3, 1).Color = Yellow
            C4.Pixels(3, 2).Color = Yellow
            C4.Pixels(3, 3).Color = Black
            C4.Pixels(3, 4).Color = Yellow
            C4.Pixels(3, 5).Color = Yellow
            C4.Pixels(3, 6).Color = Yellow
            'row 4
            C4.Pixels(4, 1).Color = Yellow
            C4.Pixels(4, 2).Color = Yellow
            C4.Pixels(4, 3).Color = Yellow
            C4.Pixels(4, 4).Color = Yellow
            C4.Pixels(4, 5).Color = Yellow
            'row 5
            C4.Pixels(5, 1).Color = Yellow
            C4.Pixels(5, 2).Color = Yellow
            C4.Pixels(5, 3).Color = Yellow
            C4.Pixels(5, 4).Color = Yellow
            C4.Pixels(5, 5).Color = Yellow
            'row 6
            C4.Pixels(6, 1).Color = Yellow
            C4.Pixels(6, 2).Color = Yellow
            C4.Pixels(6, 3).Color = Yellow
            C4.Pixels(6, 4).Color = Yellow
            'row 7
            C4.Pixels(7, 1).Color = Yellow
            C4.Pixels(7, 2).Color = Yellow
    
    'screen updating
    Application.ScreenUpdating = False
    
    'iterate cubes
    For ic = 1 To 4
        'grab next cube object
        If ic = 1 Then
            Set tCube = C1
        ElseIf ic = 2 Then
            Set tCube = C2
        ElseIf ic = 3 Then
            Set tCube = C3
        ElseIf ic = 4 Then
            Set tCube = C4
        End If
        'draw cube
        tCube.Draw
    Next ic
    
    'screen updating
    Application.ScreenUpdating = True

End Sub
        
Sub test_RangesSubtraction()

    Dim r1 As Range
    Dim r2 As Range
    Dim r3 As Range
    Dim c As Range 'iterator
    
    Set r1 = Range(Cells(1, 1), Cells(100, 100))
    Set r2 = Range(Cells(50, 50), Cells(70, 150))
    
    Set r3 = Not Intersect(r1, r2)

End Sub
