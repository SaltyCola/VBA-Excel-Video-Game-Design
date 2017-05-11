VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sprite_Map"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'=================================================================================='
'=================================== Properties ==================================='
'=================================================================================='

'Map Tile Colors
Private pColorOutline As Long 'Color of the outline of map tiles
Private pColorFill As Long 'Color of the fill of map tiles
Private pColorTrack As Long 'Color of the map track tiles (off black)
Private pColorGate As Long 'Color of the Ghost Cage Gate tile
Private pColorDots As Long 'Color of the PacDots and PowerDots

'Tiles List
Private pTiles As List_MapTiles 'List of all possible map tiles

'PacDots List
Private pDots() As Sprite_PacDots 'List of all PacDots in the level
Private pDotsCnt As Integer 'Counter for adding PacDots
Private pDotsLeft As Integer 'Number of pacdots left uneaten

'Map Seed
Private pLevelSeed() As String 'String Matrix array for loading level map

'Map Array
Private pWidthCubes As Integer 'Width of the map in Cubes
Private pHeightCubes As Integer 'Height of the map in Cubes
Private pMapArray() As Cube 'Array representing the Level Map

'''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
' Tile Color Outline
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Property Get ColorOutline() As Long
    ColorOutline = pColorOutline
End Property
Public Property Let ColorOutline(Value As Long)
    pColorOutline = Value
End Property

'''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
' Tile Color Fill
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Property Get ColorFill() As Long
    ColorFill = pColorFill
End Property
Public Property Let ColorFill(Value As Long)
    pColorFill = Value
End Property

'''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
' Track Tile Color
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Property Get ColorTrack() As Long
    ColorTrack = pColorTrack
End Property
Public Property Let ColorTrack(Value As Long)
    pColorTrack = Value
End Property

'''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
' Tile Color Ghost Gate
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Property Get ColorGate() As Long
    ColorGate = pColorGate
End Property
Public Property Let ColorGate(Value As Long)
    pColorGate = Value
End Property

'''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
' Tile Color PacDots
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Property Get ColorDots() As Long
    ColorDots = pColorDots
End Property
Public Property Let ColorDots(Value As Long)
    pColorDots = Value
End Property

'''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
' Tiles List
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Property Get Tiles() As List_MapTiles
    Set Tiles = pTiles
End Property
Public Property Let Tiles(Value As List_MapTiles)
    Set pTiles = Value
End Property

'''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
' PacDots List
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Property Get Dots(ByVal indDot As Integer) As Sprite_PacDots
    Set Dots = pDots(indDot)
End Property
Public Property Let Dots(ByVal indDot As Integer, ByRef Value As Sprite_PacDots)
    Set pDots(indDot) = Value
End Property

'''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
' Number of PacDots in List
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Property Get DotsCnt() As Integer
    DotsCnt = pDotsCnt
End Property
Public Property Let DotsCnt(Value As Integer)
    pDotsCnt = Value
End Property

'''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
' Number of PacDots Left in Level
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Property Get DotsLeft() As Integer
    DotsLeft = pDotsLeft
End Property
Public Property Let DotsLeft(Value As Integer)
    pDotsLeft = Value
End Property

'''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
' Level Seed Array (READ ONLY)
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Property Get LevelSeed(ByVal Yind As Integer, ByVal Xind As Integer) As Cube
    Set LevelSeed = pLevelSeed(Yind, Xind)
End Property
'Public Property Let LevelSeed(ByVal Yind As Integer, ByVal Xind As Integer, ByRef Value As Cube)
'    Set pLevelSeed(Yind, Xind) = Value
'End Property

'''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
' Width of Map in Cubes
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Property Get WidthCubes() As Integer
    WidthCubes = pWidthCubes
End Property
Public Property Let WidthCubes(Value As Integer)
    pWidthCubes = Value
End Property

'''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
' Height of Map in Cubes
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Property Get HeightCubes() As Integer
    HeightCubes = pHeightCubes
End Property
Public Property Let HeightCubes(Value As Integer)
    pHeightCubes = Value
End Property

'''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
' Level Map Array
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Property Get MapArray(ByVal Yind As Integer, ByVal Xind As Integer) As Cube
    Set MapArray = pMapArray(Yind, Xind)
End Property
Public Property Let MapArray(ByVal Yind As Integer, ByVal Xind As Integer, ByRef Value As Cube)
    Set pMapArray(Yind, Xind) = Value
End Property


'================================================================================='
'================================ Public Methods ================================='
'================================================================================='

Public Sub Map_Initialize()
'requires width and height to be set before calling map initialize

    'initialize Map and Seed Arrays
    ReDim pMapArray(1 To pHeightCubes, 1 To pWidthCubes)
    ReDim pLevelSeed(1 To pHeightCubes, 1 To pWidthCubes)
    
    'pass color props
    pTiles.ColorOutline = pColorOutline
    pTiles.ColorFill = pColorFill
    pTiles.ColorTrack = pColorTrack
    pTiles.ColorGate = pColorGate
    'pColorDots is not necessary to pass as PacDots are separate from map tiles

End Sub

Public Sub Seed_Load(ByVal SeedString As String)
'takes the level seed as a string arguement and creates a string array, stored in this object

    Dim iy As Integer 'integer iterator for rows
    Dim ix As Integer 'integer iterator for columns
    Dim i As Integer 'counter used for splitting seed string into an array matrix
    Dim arrString As Variant 'array matrix object for storing split seed string
    
    'split seed string into one long column array
    arrString = Split(SeedString)
    
    'initialize counter
    i = 0
    
    'organize arrString into a matrix array and apply to object's Level Seed property
    For iy = 1 To pHeightCubes
        For ix = 1 To pWidthCubes
            pLevelSeed(iy, ix) = arrString(i)
            i = i + 1 'increment counter
        Next ix
    Next iy

End Sub

Public Sub Map_Load()
'map must be initialized first and level seed given

    Dim iy As Integer 'integer iterator for rows
    Dim ix As Integer 'integer iterator for columns
    Dim iy2 As Integer 'integer iterator for rows (secondary)
    Dim ix2 As Integer 'integer iterator for columns (secondary)
    Dim cx As Integer 'integer for applying correct cube coordinates
    Dim cy As Integer 'integer for applying correct cube coordinates
    Dim tCube As Cube 'temporary cube object as mediary between level seed array and map array
    Dim tTile As Cube 'tile cube from tiles list
    
    'initialize PacDots List
    ReDim pDots(1 To 1) As Sprite_PacDots
    pDotsCnt = 0
    
    'initialize x and y cube position counters
    cx = 17 'leaves hidden rows on left for character sprites to go through
           ' the tunnel without causing the game to crash due to negative x values.
    cy = 25 'leaves space above for game headers
    
    'create a cube object in MapArray for every tile code in LevelSeed
    For iy = 1 To pHeightCubes
        For ix = 1 To pWidthCubes
            'grab tile cube
            Set tTile = pTiles.Tile(pLevelSeed(iy, ix))
            'initialize new cube
            Set tCube = New Cube
            tCube.Xcol = 1
            tCube.Yrow = 1
            tCube.SetPixels
            'COPY pixel colors from seed tile code
            For iy2 = 1 To 8
                For ix2 = 1 To 8
                    tCube.Pixels(iy2, ix2).Color = tTile.Pixels(iy2, ix2).Color
                    tCube.DotType = tTile.DotType
                Next ix2
            Next iy2
            'update cube's position
            tCube.Xcol = cx
            tCube.Yrow = cy
            'update cube pixel positions
            tCube.UpdatePixels
            'add cube to array
            Set pMapArray(iy, ix) = tCube
            'Add PacDot to Dots list
            Me.AddPacDot tTile.DotType, cx, cy
            'increment counters
            If cx = 233 Then 'end of row, start next row
                cx = 17 'accounts for 2 hidden cubes on left of level map
                cy = cy + 8
            Else
                cx = cx + 8
            End If
        Next ix
    Next iy

End Sub

Public Sub Map_Draw()
'map must be initialized, then seed loaded, then map loaded before this method is used

    Dim iy As Integer 'integer iterator for rows
    Dim ix As Integer 'integer iterator for columns
    Dim iDots As Integer 'counter for which PacDot object is currently being looked at
    Dim lBar As frm_LoadingBar 'level loading bar
    Dim iLoad1 As Integer 'loading bar counter
    Dim iLoad2 As Integer 'loading bar counter
    
    'initialize loading bar
    Set lBar = New frm_LoadingBar
    iLoad1 = 0
    iLoad2 = 0
    
    'initialize dots counter
    iDots = 0
    
    'iterate cubes in map sprite
    For iy = 1 To 31
        For ix = 1 To 28
            'send pacdot color to track tiles with pacdots
                If pMapArray(iy, ix).DotType = 1 Then: pMapArray(iy, ix).DotColor = pColorDots
            'draw tile cube to map ranges
                pMapArray(iy, ix).Map_Draw
            'update loading bar
                iLoad1 = iLoad1 + 1
                If iLoad1 = 4 Then
                    iLoad2 = iLoad2 + 1
                    lBar.UpdateLoadingBar "Level Loading...", iLoad2, 217 '217 = (31*28)/4 (divide by 4 to reduce amount of times called)
                    iLoad1 = 0
                End If
        Next ix
    Next iy
    
    'Erase Loading Bar
    Set lBar = Nothing

End Sub

Public Sub AddPacDot(ByVal indDotType As Integer, ByVal indXcol As Integer, ByVal indYrow As Integer)

    If indDotType <> 0 Then
    
        'increment dots count
        pDotsCnt = pDotsCnt + 1
        
        'Update Dots Left
        pDotsLeft = pDotsCnt
        
        'add array entry
        ReDim Preserve pDots(1 To pDotsCnt) As Sprite_PacDots
        
        'create new PacDot Sprite
        Set pDots(pDotsCnt) = New Sprite_PacDots
        
        'initialize PacDot properties
        pDots(pDotsCnt).Xcol = indXcol
        pDots(pDotsCnt).Yrow = indYrow
        pDots(pDotsCnt).Color = Me.ColorDots
        pDots(pDotsCnt).nAnim = indDotType
        If pDots(pDotsCnt).nAnim = 2 Then: pDots(pDotsCnt).PowerDot = True
        
        'initialize cubes' properties
        pDots(pDotsCnt).SetCubes
        pDots(pDotsCnt).SetColorRange
    
    End If

End Sub


'================================================================================='
'=============================== Private Methods ================================='
'================================================================================='

Private Sub Class_Initialize()

    'initialize tile list
    Set pTiles = New List_MapTiles

End Sub
