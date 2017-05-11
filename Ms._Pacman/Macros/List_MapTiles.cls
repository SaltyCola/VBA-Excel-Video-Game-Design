VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "List_MapTiles"
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

'2-Dir Track Tiles
    'corner tiles
    Private pTL_Track As Cube 'top left corner                          (1)
    Private pTR_Track As Cube 'top right corner                         (2)
    Private pBR_Track As Cube 'bottom right corner                      (3)
    Private pBL_Track As Cube 'bottom left corner                       (4)
    'straight tiles
    Private pL_Track As Cube 'left-side tile                            (5)
    Private pT_Track As Cube 'top-side tile                             (6)
'Multi-Dir Track Tiles
    'T-tiles
    Private pLT_Track As Cube 'left T track                             (1)
    Private pTT_Track As Cube 'top T track                              (2)
    Private pRT_Track As Cube 'right T track                            (3)
    Private pBT_Track As Cube 'bottom T track                           (4)
    'cross tile
    Private pC_Track As Cube '4 dir track                               (5)

'Outter Wall Tiles
    'corner tiles
    Private pTL_Outter As Cube 'top left corner                         (1)
    Private pTR_Outter As Cube 'top right corner                        (2)
    Private pBR_Outter As Cube 'bottom right corner                     (3)
    Private pBL_Outter As Cube 'bottom left corner                      (4)
    'straight tiles
    Private pL_Outter As Cube 'left-side tile                           (5)
    Private pT_Outter As Cube 'top-side tile                            (6)
    Private pR_Outter As Cube 'right-side tile                          (7)
    Private pB_Outter As Cube 'bottom-side tile                         (8)

'Inner Wall Tiles
    'corner tiles
    Private pTL_Inner As Cube 'top left corner                          (1)
    Private pTR_Inner As Cube 'top right corner                         (2)
    Private pBR_Inner As Cube 'bottom right corner                      (3)
    Private pBL_Inner As Cube 'bottom left corner                       (4)
    'straight tiles
    Private pL_Inner As Cube 'left-side tile                            (5)
    Private pT_Inner As Cube 'top-side tile                             (6)
    Private pR_Inner As Cube 'right-side tile                           (7)
    Private pB_Inner As Cube 'bottom-side tile                          (8)
    'inverted corner tiles
    Private pTL_InnerInv As Cube 'top left corner INVERTED              (1)
    Private pTR_InnerInv As Cube 'top right corner INVERTED             (2)
    Private pBR_InnerInv As Cube 'bottom right corner INVERTED          (3)
    Private pBL_InnerInv As Cube 'bottom left corner INVERTED           (4)
    'inverted corner tiles OUTER EDGE (have outline)
    Private pTL_InnerInvOL As Cube 'top left corner INVERTED OUTLINE    (5)
    Private pTR_InnerInvOL As Cube 'top right corner INVERTED OUTLINE   (6)
    Private pBR_InnerInvOL As Cube 'bottom right corner INVERTED OUTLINE(7)
    Private pBL_InnerInvOL As Cube 'bottom left corner INVERTED OUTLINE (8)

'Filled Cube Tiles
    Private pF_Blank As Cube 'blank tile (all black)                    (0)
    Private pF_Inner As Cube 'full fill-color tile                      (1)

'Ghost Cage Tiles
    'corner tiles
    Private pTL_GhostCage As Cube 'top left corner                      (1)
    Private pTR_GhostCage As Cube 'top right corner                     (2)
    Private pBR_GhostCage As Cube 'bottom right corner                  (3)
    Private pBL_GhostCage As Cube 'bottom left corner                   (4)
    'straight tiles
    Private pL_GhostCage As Cube 'left-side tile (same as R_Outter)     (5)
    Private pT_GhostCage As Cube 'top-side tile (same as B_Outter)      (6)
    Private pR_GhostCage As Cube 'right-side tile (same as L_Outter)    (7)
    Private pB_GhostCage As Cube 'bottom-side tile (same as T_Outter)   (8)
    'cage gate tile
    Private pGL_GhostCage As Cube 'ghost cage's Left gate tile          (1)
    Private pGR_GhostCage As Cube 'ghost cage's Right gate tile         (2)

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
' Track Tiles
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Property Get TL_Track() As Cube
    Set TL_Track = pTL_Track
End Property
Public Property Let TL_Track(Value As Cube)
    Set pTL_Track = Value
End Property

Public Property Get TR_Track() As Cube
    Set TR_Track = pTR_Track
End Property
Public Property Let TR_Track(Value As Cube)
    Set pTR_Track = Value
End Property

Public Property Get BR_Track() As Cube
    Set BR_Track = pBR_Track
End Property
Public Property Let BR_Track(Value As Cube)
    Set pBR_Track = Value
End Property

Public Property Get BL_Track() As Cube
    Set BL_Track = pBL_Track
End Property
Public Property Let BL_Track(Value As Cube)
    Set pBL_Track = Value
End Property

Public Property Get L_Track() As Cube
    Set L_Track = pL_Track
End Property
Public Property Let L_Track(Value As Cube)
    Set pL_Track = Value
End Property

Public Property Get T_Track() As Cube
    Set T_Track = pT_Track
End Property
Public Property Let T_Track(Value As Cube)
    Set pT_Track = Value
End Property

Public Property Get LT_Track() As Cube
    Set LT_Track = pLT_Track
End Property
Public Property Let LT_Track(Value As Cube)
    Set pLT_Track = Value
End Property

Public Property Get TT_Track() As Cube
    Set TT_Track = pTT_Track
End Property
Public Property Let TT_Track(Value As Cube)
    Set pTT_Track = Value
End Property

Public Property Get RT_Track() As Cube
    Set RT_Track = pRT_Track
End Property
Public Property Let RT_Track(Value As Cube)
    Set pRT_Track = Value
End Property

Public Property Get BT_Track() As Cube
    Set BT_Track = pBT_Track
End Property
Public Property Let BT_Track(Value As Cube)
    Set pBT_Track = Value
End Property

Public Property Get C_Track() As Cube
    Set C_Track = pC_Track
End Property
Public Property Let C_Track(Value As Cube)
    Set pC_Track = Value
End Property

'''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
' Outter Wall Tiles
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''

'Corner Tiles

Public Property Get TL_Outter() As Cube
    Set TL_Outter = pTL_Outter
End Property
Public Property Let TL_Outter(Value As Cube)
    Set pTL_Outter = Value
End Property

Public Property Get TR_Outter() As Cube
    Set TR_Outter = pTR_Outter
End Property
Public Property Let TR_Outter(Value As Cube)
    Set pTR_Outter = Value
End Property

Public Property Get BL_Outter() As Cube
    Set BL_Outter = pBL_Outter
End Property
Public Property Let BL_Outter(Value As Cube)
    Set pBL_Outter = Value
End Property

Public Property Get BR_Outter() As Cube
    Set BR_Outter = pBR_Outter
End Property
Public Property Let BR_Outter(Value As Cube)
    Set pBR_Outter = Value
End Property

'Straight Tiles

Public Property Get L_Outter() As Cube
    Set L_Outter = pL_Outter
End Property
Public Property Let L_Outter(Value As Cube)
    Set pL_Outter = Value
End Property

Public Property Get R_Outter() As Cube
    Set R_Outter = pR_Outter
End Property
Public Property Let R_Outter(Value As Cube)
    Set pR_Outter = Value
End Property

Public Property Get T_Outter() As Cube
    Set T_Outter = pT_Outter
End Property
Public Property Let T_Outter(Value As Cube)
    Set pT_Outter = Value
End Property

Public Property Get B_Outter() As Cube
    Set B_Outter = pB_Outter
End Property
Public Property Let B_Outter(Value As Cube)
    Set pB_Outter = Value
End Property

'''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
' Inner Wall Tiles
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''

'Corner Tiles

Public Property Get TL_Inner() As Cube
    Set TL_Inner = pTL_Inner
End Property
Public Property Let TL_Inner(Value As Cube)
    Set pTL_Inner = Value
End Property

Public Property Get TR_Inner() As Cube
    Set TR_Inner = pTR_Inner
End Property
Public Property Let TR_Inner(Value As Cube)
    Set pTR_Inner = Value
End Property

Public Property Get BL_Inner() As Cube
    Set BL_Inner = pBL_Inner
End Property
Public Property Let BL_Inner(Value As Cube)
    Set pBL_Inner = Value
End Property

Public Property Get BR_Inner() As Cube
    Set BR_Inner = pBR_Inner
End Property
Public Property Let BR_Inner(Value As Cube)
    Set pBR_Inner = Value
End Property

'Straight Tiles

Public Property Get L_Inner() As Cube
    Set L_Inner = pL_Inner
End Property
Public Property Let L_Inner(Value As Cube)
    Set pL_Inner = Value
End Property

Public Property Get R_Inner() As Cube
    Set R_Inner = pR_Inner
End Property
Public Property Let R_Inner(Value As Cube)
    Set pR_Inner = Value
End Property

Public Property Get T_Inner() As Cube
    Set T_Inner = pT_Inner
End Property
Public Property Let T_Inner(Value As Cube)
    Set pT_Inner = Value
End Property

Public Property Get B_Inner() As Cube
    Set B_Inner = pB_Inner
End Property
Public Property Let B_Inner(Value As Cube)
    Set pB_Inner = Value
End Property

'Inverted Corner Tiles

Public Property Get TL_InnerInv() As Cube
    Set TL_InnerInv = pTL_InnerInv
End Property
Public Property Let TL_InnerInv(Value As Cube)
    Set pTL_InnerInv = Value
End Property

Public Property Get TR_InnerInv() As Cube
    Set TR_InnerInv = pTR_InnerInv
End Property
Public Property Let TR_InnerInv(Value As Cube)
    Set pTR_InnerInv = Value
End Property

Public Property Get BL_InnerInv() As Cube
    Set BL_InnerInv = pBL_InnerInv
End Property
Public Property Let BL_InnerInv(Value As Cube)
    Set pBL_InnerInv = Value
End Property

Public Property Get BR_InnerInv() As Cube
    Set BR_InnerInv = pBR_InnerInv
End Property
Public Property Let BR_InnerInv(Value As Cube)
    Set pBR_InnerInv = Value
End Property

Public Property Get TL_InnerInvOL() As Cube
    Set TL_InnerInvOL = pTL_InnerInvOL
End Property
Public Property Let TL_InnerInvOL(Value As Cube)
    Set pTL_InnerInvOL = Value
End Property

Public Property Get TR_InnerInvOL() As Cube
    Set TR_InnerInvOL = pTR_InnerInvOL
End Property
Public Property Let TR_InnerInvOL(Value As Cube)
    Set pTR_InnerInvOL = Value
End Property

Public Property Get BL_InnerInvOL() As Cube
    Set BL_InnerInvOL = pBL_InnerInvOL
End Property
Public Property Let BL_InnerInvOL(Value As Cube)
    Set pBL_InnerInvOL = Value
End Property

Public Property Get BR_InnerInvOL() As Cube
    Set BR_InnerInvOL = pBR_InnerInvOL
End Property
Public Property Let BR_InnerInvOL(Value As Cube)
    Set pBR_InnerInvOL = Value
End Property

'''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
'Filled Tiles
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Property Get F_Blank() As Cube
    Set F_Blank = pF_Blank
End Property
Public Property Let F_Blank(Value As Cube)
    Set pF_Blank = Value
End Property

Public Property Get F_Inner() As Cube
    Set F_Inner = pF_Inner
End Property
Public Property Let F_Inner(Value As Cube)
    Set pF_Inner = Value
End Property

'''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
' Ghost Cage Tiles
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''

'Corner Tiles

Public Property Get TL_GhostCage() As Cube
    Set TL_GhostCage = pTL_GhostCage
End Property
Public Property Let TL_GhostCage(Value As Cube)
    Set pTL_GhostCage = Value
End Property

Public Property Get TR_GhostCage() As Cube
    Set TR_GhostCage = pTR_GhostCage
End Property
Public Property Let TR_GhostCage(Value As Cube)
    Set pTR_GhostCage = Value
End Property

Public Property Get BL_GhostCage() As Cube
    Set BL_GhostCage = pBL_GhostCage
End Property
Public Property Let BL_GhostCage(Value As Cube)
    Set pBL_GhostCage = Value
End Property

Public Property Get BR_GhostCage() As Cube
    Set BR_GhostCage = pBR_GhostCage
End Property
Public Property Let BR_GhostCage(Value As Cube)
    Set pBR_GhostCage = Value
End Property

'Straight Tiles

Public Property Get L_GhostCage() As Cube
    Set L_GhostCage = pL_GhostCage
End Property
Public Property Let L_GhostCage(Value As Cube)
    Set pL_GhostCage = Value
End Property

Public Property Get R_GhostCage() As Cube
    Set R_GhostCage = pR_GhostCage
End Property
Public Property Let R_GhostCage(Value As Cube)
    Set pR_GhostCage = Value
End Property

Public Property Get T_GhostCage() As Cube
    Set T_GhostCage = pT_GhostCage
End Property
Public Property Let T_GhostCage(Value As Cube)
    Set pT_GhostCage = Value
End Property

Public Property Get B_GhostCage() As Cube
    Set B_GhostCage = pB_GhostCage
End Property
Public Property Let B_GhostCage(Value As Cube)
    Set pB_GhostCage = Value
End Property

'Gate Tiles

Public Property Get GL_GhostCage() As Cube
    Set GL_GhostCage = pGL_GhostCage
End Property
Public Property Let GL_GhostCage(Value As Cube)
    Set pGL_GhostCage = Value
End Property

Public Property Get GR_GhostCage() As Cube
    Set GR_GhostCage = pGR_GhostCage
End Property
Public Property Let GR_GhostCage(Value As Cube)
    Set pGR_GhostCage = Value
End Property


'================================================================================='
'================================ Public Methods ================================='
'================================================================================='

Public Sub Initialize_Tiles()

    'call subdivided intializers
    Call Initialize_TrackTiles
    Call Initialize_OutterTiles
    Call Initialize_InnerTiles 'filled tiles in with this sub
    Call Initialize_GhostCageTiles

End Sub

Public Function Tile(ByVal TileCode As String) As Cube

    Dim tCube As Cube 'temporary cube object for returning correct object
    
    'Grab corresponding Map Tile
    If left(TileCode, (Len(TileCode) - 1)) = "t1" Then
        Set tCube = pTL_Track
    ElseIf left(TileCode, (Len(TileCode) - 1)) = "t2" Then
        Set tCube = pTR_Track
    ElseIf left(TileCode, (Len(TileCode) - 1)) = "t3" Then
        Set tCube = pBR_Track
    ElseIf left(TileCode, (Len(TileCode) - 1)) = "t4" Then
        Set tCube = pBL_Track
    ElseIf left(TileCode, (Len(TileCode) - 1)) = "t5" Then
        Set tCube = pL_Track
    ElseIf left(TileCode, (Len(TileCode) - 1)) = "t6" Then
        Set tCube = pT_Track
    ElseIf left(TileCode, (Len(TileCode) - 1)) = "x1" Then
        Set tCube = pLT_Track
    ElseIf left(TileCode, (Len(TileCode) - 1)) = "x2" Then
        Set tCube = pTT_Track
    ElseIf left(TileCode, (Len(TileCode) - 1)) = "x3" Then
        Set tCube = pRT_Track
    ElseIf left(TileCode, (Len(TileCode) - 1)) = "x4" Then
        Set tCube = pBT_Track
    ElseIf left(TileCode, (Len(TileCode) - 1)) = "x5" Then
        Set tCube = pC_Track
    ElseIf left(TileCode, (Len(TileCode) - 1)) = "o1" Then
        Set tCube = pTL_Outter
    ElseIf left(TileCode, (Len(TileCode) - 1)) = "o2" Then
        Set tCube = pTR_Outter
    ElseIf left(TileCode, (Len(TileCode) - 1)) = "o3" Then
        Set tCube = pBR_Outter
    ElseIf left(TileCode, (Len(TileCode) - 1)) = "o4" Then
        Set tCube = pBL_Outter
    ElseIf left(TileCode, (Len(TileCode) - 1)) = "o5" Then
        Set tCube = pL_Outter
    ElseIf left(TileCode, (Len(TileCode) - 1)) = "o6" Then
        Set tCube = pT_Outter
    ElseIf left(TileCode, (Len(TileCode) - 1)) = "o7" Then
        Set tCube = pR_Outter
    ElseIf left(TileCode, (Len(TileCode) - 1)) = "o8" Then
        Set tCube = pB_Outter
    ElseIf left(TileCode, (Len(TileCode) - 1)) = "i1" Then
        Set tCube = pTL_Inner
    ElseIf left(TileCode, (Len(TileCode) - 1)) = "i2" Then
        Set tCube = pTR_Inner
    ElseIf left(TileCode, (Len(TileCode) - 1)) = "i3" Then
        Set tCube = pBR_Inner
    ElseIf left(TileCode, (Len(TileCode) - 1)) = "i4" Then
        Set tCube = pBL_Inner
    ElseIf left(TileCode, (Len(TileCode) - 1)) = "i5" Then
        Set tCube = pL_Inner
    ElseIf left(TileCode, (Len(TileCode) - 1)) = "i6" Then
        Set tCube = pT_Inner
    ElseIf left(TileCode, (Len(TileCode) - 1)) = "i7" Then
        Set tCube = pR_Inner
    ElseIf left(TileCode, (Len(TileCode) - 1)) = "i8" Then
        Set tCube = pB_Inner
    ElseIf left(TileCode, (Len(TileCode) - 1)) = "v1" Then
        Set tCube = pTL_InnerInv
    ElseIf left(TileCode, (Len(TileCode) - 1)) = "v2" Then
        Set tCube = pTR_InnerInv
    ElseIf left(TileCode, (Len(TileCode) - 1)) = "v3" Then
        Set tCube = pBR_InnerInv
    ElseIf left(TileCode, (Len(TileCode) - 1)) = "v4" Then
        Set tCube = pBL_InnerInv
    ElseIf left(TileCode, (Len(TileCode) - 1)) = "v5" Then
        Set tCube = pTL_InnerInvOL
    ElseIf left(TileCode, (Len(TileCode) - 1)) = "v6" Then
        Set tCube = pTR_InnerInvOL
    ElseIf left(TileCode, (Len(TileCode) - 1)) = "v7" Then
        Set tCube = pBR_InnerInvOL
    ElseIf left(TileCode, (Len(TileCode) - 1)) = "v8" Then
        Set tCube = pBL_InnerInvOL
    ElseIf left(TileCode, (Len(TileCode) - 1)) = "f0" Then
        Set tCube = pF_Blank
    ElseIf left(TileCode, (Len(TileCode) - 1)) = "f1" Then
        Set tCube = pF_Inner
    ElseIf left(TileCode, (Len(TileCode) - 1)) = "c1" Then
        Set tCube = pTL_GhostCage
    ElseIf left(TileCode, (Len(TileCode) - 1)) = "c2" Then
        Set tCube = pTR_GhostCage
    ElseIf left(TileCode, (Len(TileCode) - 1)) = "c3" Then
        Set tCube = pBR_GhostCage
    ElseIf left(TileCode, (Len(TileCode) - 1)) = "c4" Then
        Set tCube = pBL_GhostCage
    ElseIf left(TileCode, (Len(TileCode) - 1)) = "c5" Then
        Set tCube = pL_GhostCage
    ElseIf left(TileCode, (Len(TileCode) - 1)) = "c6" Then
        Set tCube = pT_GhostCage
    ElseIf left(TileCode, (Len(TileCode) - 1)) = "c7" Then
        Set tCube = pR_GhostCage
    ElseIf left(TileCode, (Len(TileCode) - 1)) = "c8" Then
        Set tCube = pB_GhostCage
    ElseIf left(TileCode, (Len(TileCode) - 1)) = "g1" Then
        Set tCube = pGL_GhostCage
    ElseIf left(TileCode, (Len(TileCode) - 1)) = "g2" Then
        Set tCube = pGR_GhostCage
    End If
    
    'Give Tile a PacDot or PowerDot?
    If Right(TileCode, 1) = "o" Then 'none
        tCube.DotType = 0
    ElseIf Right(TileCode, 1) = "p" Then 'PacDot
        tCube.DotType = 1
    ElseIf Right(TileCode, 1) = "P" Then 'PowerDot
        tCube.DotType = 2
    End If
    
    'function return
    Set Tile = tCube

End Function


'================================================================================='
'================================ Private Methods ================================'
'================================================================================='

Private Sub Class_Initialize()

    Dim i As Integer 'integer iterator
    
    'initialize cube objects
        'track
        Set pTL_Track = New Cube
        Set pTR_Track = New Cube
        Set pBR_Track = New Cube
        Set pBL_Track = New Cube
        Set pL_Track = New Cube
        Set pT_Track = New Cube
        Set pLT_Track = New Cube
        Set pTT_Track = New Cube
        Set pRT_Track = New Cube
        Set pBT_Track = New Cube
        Set pC_Track = New Cube
        'outter
        Set pTL_Outter = New Cube
        Set pTR_Outter = New Cube
        Set pBL_Outter = New Cube
        Set pBR_Outter = New Cube
        Set pL_Outter = New Cube
        Set pR_Outter = New Cube
        Set pT_Outter = New Cube
        Set pB_Outter = New Cube
        'inner
        Set pTL_Inner = New Cube
        Set pTR_Inner = New Cube
        Set pBL_Inner = New Cube
        Set pBR_Inner = New Cube
        Set pL_Inner = New Cube
        Set pR_Inner = New Cube
        Set pT_Inner = New Cube
        Set pB_Inner = New Cube
        'filled tiles
        Set pF_Blank = New Cube
        Set pF_Inner = New Cube
        'inner inverted
        Set pTL_InnerInv = New Cube
        Set pTR_InnerInv = New Cube
        Set pBL_InnerInv = New Cube
        Set pBR_InnerInv = New Cube
        'inner inverted outline
        Set pTL_InnerInvOL = New Cube
        Set pTR_InnerInvOL = New Cube
        Set pBL_InnerInvOL = New Cube
        Set pBR_InnerInvOL = New Cube
        'ghost cage
        Set pTL_GhostCage = New Cube
        Set pTR_GhostCage = New Cube
        Set pBL_GhostCage = New Cube
        Set pBR_GhostCage = New Cube
        Set pL_GhostCage = New Cube
        Set pR_GhostCage = New Cube
        Set pT_GhostCage = New Cube
        Set pB_GhostCage = New Cube
        'ghost gate
        Set pGL_GhostCage = New Cube
        Set pGR_GhostCage = New Cube
        
    'initialize pixel object properties
        'track tiles 2-dir
        For i = 1 To 6
            Me.Tile("t" & i & "o").Xcol = 1
            Me.Tile("t" & i & "o").Yrow = 1
            Me.Tile("t" & i & "o").SetPixels
        Next i
        'track tiles multi-dir (x = cross/intersection tracks)
        For i = 1 To 5
            Me.Tile("x" & i & "o").Xcol = 1
            Me.Tile("x" & i & "o").Yrow = 1
            Me.Tile("x" & i & "o").SetPixels
        Next i
        'outter
        For i = 1 To 8
            Me.Tile("o" & i & "o").Xcol = 1
            Me.Tile("o" & i & "o").Yrow = 1
            Me.Tile("o" & i & "o").SetPixels
        Next i
        'inner
        For i = 1 To 8
            Me.Tile("i" & i & "o").Xcol = 1
            Me.Tile("i" & i & "o").Yrow = 1
            Me.Tile("i" & i & "o").SetPixels
        Next i
        'blank tile
            Me.Tile("f0o").Xcol = 1
            Me.Tile("f0o").Yrow = 1
            Me.Tile("f0o").SetPixels
        'filled tile
            Me.Tile("f1o").Xcol = 1
            Me.Tile("f1o").Yrow = 1
            Me.Tile("f1o").SetPixels
        'inner inverted
        For i = 1 To 8
            Me.Tile("v" & i & "o").Xcol = 1
            Me.Tile("v" & i & "o").Yrow = 1
            Me.Tile("v" & i & "o").SetPixels
        Next i
        'ghost cage
        For i = 1 To 8
            Me.Tile("c" & i & "o").Xcol = 1
            Me.Tile("c" & i & "o").Yrow = 1
            Me.Tile("c" & i & "o").SetPixels
        Next i
        'ghost gate
        For i = 1 To 2
            Me.Tile("g" & i & "o").Xcol = 1
            Me.Tile("g" & i & "o").Yrow = 1
            Me.Tile("g" & i & "o").SetPixels
        Next i

End Sub

Private Sub Initialize_TrackTiles()

    Dim ix As Integer 'integer iterator for x values of pixel arrays
    Dim iy As Integer 'integer iterator for y values of pixel arrays
    
    'CORNER TILE
        'Top-Left
            'row 4
            pTL_Track.Pixels(4, 4).Color = pColorTrack
            pTL_Track.Pixels(4, 5).Color = pColorTrack
            pTL_Track.Pixels(4, 6).Color = pColorTrack
            pTL_Track.Pixels(4, 7).Color = pColorTrack
            pTL_Track.Pixels(4, 8).Color = pColorTrack
            'row 5
            pTL_Track.Pixels(5, 4).Color = pColorTrack
            pTL_Track.Pixels(5, 5).Color = pColorTrack
            pTL_Track.Pixels(5, 6).Color = pColorTrack
            pTL_Track.Pixels(5, 7).Color = pColorTrack
            pTL_Track.Pixels(5, 8).Color = pColorTrack
            'row 6
            pTL_Track.Pixels(6, 4).Color = pColorTrack
            pTL_Track.Pixels(6, 5).Color = pColorTrack
            'row 7
            pTL_Track.Pixels(7, 4).Color = pColorTrack
            pTL_Track.Pixels(7, 5).Color = pColorTrack
            'row 8
            pTL_Track.Pixels(8, 4).Color = pColorTrack
            pTL_Track.Pixels(8, 5).Color = pColorTrack
        'Top-Right
            For iy = 1 To 8
                For ix = 1 To 8
                    pTR_Track.Pixels(iy, ix).Color = pTL_Track.Pixels(iy, (9 - ix)).Color
                Next ix
            Next iy
        'Bottom-Left
            For iy = 1 To 8
                For ix = 1 To 8
                    pBL_Track.Pixels(iy, ix).Color = pTL_Track.Pixels((9 - iy), ix).Color
                Next ix
            Next iy
        'Bottom-Right
            For iy = 1 To 8
                For ix = 1 To 8
                    pBR_Track.Pixels(iy, ix).Color = pTL_Track.Pixels((9 - iy), (9 - ix)).Color
                Next ix
            Next iy
    
    'STRAIGHT TILES
        'Left
            For iy = 1 To 8
                pL_Track.Pixels(iy, 4).Color = pColorTrack
                pL_Track.Pixels(iy, 5).Color = pColorTrack
            Next iy
        'Top
            For ix = 1 To 8
                pT_Track.Pixels(4, ix).Color = pColorTrack
                pT_Track.Pixels(5, ix).Color = pColorTrack
            Next ix
    
    'T-TRACK TILES
        'Left
            For iy = 1 To 8
                pLT_Track.Pixels(iy, 4).Color = pColorTrack
                pLT_Track.Pixels(iy, 5).Color = pColorTrack
            Next iy
            For ix = 1 To 3
                pLT_Track.Pixels(4, ix).Color = pColorTrack
                pLT_Track.Pixels(5, ix).Color = pColorTrack
            Next ix
        'Right
            For iy = 1 To 8
                For ix = 1 To 8
                    pRT_Track.Pixels(iy, ix).Color = pLT_Track.Pixels(iy, (9 - ix)).Color
                Next ix
            Next iy
        'Top
            For iy = 1 To 3
                pTT_Track.Pixels(iy, 4).Color = pColorTrack
                pTT_Track.Pixels(iy, 5).Color = pColorTrack
            Next iy
            For ix = 1 To 8
                pTT_Track.Pixels(4, ix).Color = pColorTrack
                pTT_Track.Pixels(5, ix).Color = pColorTrack
            Next ix
        'Bottom
            For iy = 1 To 8
                For ix = 1 To 8
                    pBT_Track.Pixels(iy, ix).Color = pTT_Track.Pixels((9 - iy), ix).Color
                Next ix
            Next iy
    'CROSS TRACK TILE
        'Cross (4-dir)
            For iy = 1 To 8
                pC_Track.Pixels(iy, 4).Color = pColorTrack
                pC_Track.Pixels(iy, 5).Color = pColorTrack
            Next iy
            For ix = 1 To 8
                pC_Track.Pixels(4, ix).Color = pColorTrack
                pC_Track.Pixels(5, ix).Color = pColorTrack
            Next ix

End Sub

Private Sub Initialize_OutterTiles()

    Dim ix As Integer 'integer iterator for x values of pixel arrays
    Dim iy As Integer 'integer iterator for y values of pixel arrays
    
    'CORNER TILES
        'Top-Left
            'row 1
            pTL_Outter.Pixels(1, 5).Color = pColorOutline
            pTL_Outter.Pixels(1, 6).Color = pColorOutline
            pTL_Outter.Pixels(1, 7).Color = pColorOutline
            pTL_Outter.Pixels(1, 8).Color = pColorOutline
            'row 2
            pTL_Outter.Pixels(2, 3).Color = pColorOutline
            pTL_Outter.Pixels(2, 4).Color = pColorOutline
            pTL_Outter.Pixels(2, 5).Color = pColorFill
            pTL_Outter.Pixels(2, 6).Color = pColorFill
            pTL_Outter.Pixels(2, 7).Color = pColorFill
            pTL_Outter.Pixels(2, 8).Color = pColorFill
            'row 3
            pTL_Outter.Pixels(3, 2).Color = pColorOutline
            pTL_Outter.Pixels(3, 3).Color = pColorFill
            pTL_Outter.Pixels(3, 4).Color = pColorFill
            pTL_Outter.Pixels(3, 5).Color = pColorFill
            pTL_Outter.Pixels(3, 6).Color = pColorFill
            pTL_Outter.Pixels(3, 7).Color = pColorFill
            pTL_Outter.Pixels(3, 8).Color = pColorFill
            'row 4
            pTL_Outter.Pixels(4, 2).Color = pColorOutline
            pTL_Outter.Pixels(4, 3).Color = pColorFill
            pTL_Outter.Pixels(4, 4).Color = pColorFill
            pTL_Outter.Pixels(4, 5).Color = pColorFill
            pTL_Outter.Pixels(4, 6).Color = pColorOutline
            pTL_Outter.Pixels(4, 7).Color = pColorOutline
            pTL_Outter.Pixels(4, 8).Color = pColorOutline
            'row 5
            pTL_Outter.Pixels(5, 1).Color = pColorOutline
            pTL_Outter.Pixels(5, 2).Color = pColorFill
            pTL_Outter.Pixels(5, 3).Color = pColorFill
            pTL_Outter.Pixels(5, 4).Color = pColorFill
            pTL_Outter.Pixels(5, 5).Color = pColorOutline
            'row 6
            pTL_Outter.Pixels(6, 1).Color = pColorOutline
            pTL_Outter.Pixels(6, 2).Color = pColorFill
            pTL_Outter.Pixels(6, 3).Color = pColorFill
            pTL_Outter.Pixels(6, 4).Color = pColorOutline
            'row 7
            pTL_Outter.Pixels(7, 1).Color = pColorOutline
            pTL_Outter.Pixels(7, 2).Color = pColorFill
            pTL_Outter.Pixels(7, 3).Color = pColorFill
            pTL_Outter.Pixels(7, 4).Color = pColorOutline
            'row 8
            pTL_Outter.Pixels(8, 1).Color = pColorOutline
            pTL_Outter.Pixels(8, 2).Color = pColorFill
            pTL_Outter.Pixels(8, 3).Color = pColorFill
            pTL_Outter.Pixels(8, 4).Color = pColorOutline
        'Top-Right
            For iy = 1 To 8
                For ix = 1 To 8
                    pTR_Outter.Pixels(iy, ix).Color = pTL_Outter.Pixels(iy, (9 - ix)).Color
                Next ix
            Next iy
        'Bottom-Left
            For iy = 1 To 8
                For ix = 1 To 8
                    pBL_Outter.Pixels(iy, ix).Color = pTL_Outter.Pixels((9 - iy), ix).Color
                Next ix
            Next iy
        'Bottom-Right
            For iy = 1 To 8
                For ix = 1 To 8
                    pBR_Outter.Pixels(iy, ix).Color = pTL_Outter.Pixels((9 - iy), (9 - ix)).Color
                Next ix
            Next iy
    
    'STRAIGHT TILES
        'Left
            For iy = 1 To 8
                pL_Outter.Pixels(iy, 1).Color = pColorOutline
                pL_Outter.Pixels(iy, 2).Color = pColorFill
                pL_Outter.Pixels(iy, 3).Color = pColorFill
                pL_Outter.Pixels(iy, 4).Color = pColorOutline
            Next iy
        'Right
            For iy = 1 To 8
                For ix = 1 To 8
                    pR_Outter.Pixels(iy, ix).Color = pL_Outter.Pixels(iy, (9 - ix)).Color
                Next ix
            Next iy
        'Top
            For ix = 1 To 8
                pT_Outter.Pixels(1, ix).Color = pColorOutline
                pT_Outter.Pixels(2, ix).Color = pColorFill
                pT_Outter.Pixels(3, ix).Color = pColorFill
                pT_Outter.Pixels(4, ix).Color = pColorOutline
            Next ix
        'Bottom
            For iy = 1 To 8
                For ix = 1 To 8
                    pB_Outter.Pixels(iy, ix).Color = pT_Outter.Pixels((9 - iy), ix).Color
                Next ix
            Next iy

End Sub

Private Sub Initialize_InnerTiles()

    'CORNER TILES
        'Top-Left
            'row 5
            pTL_Inner.Pixels(5, 7).Color = pColorOutline
            pTL_Inner.Pixels(5, 8).Color = pColorOutline
            'row 6
            pTL_Inner.Pixels(6, 6).Color = pColorOutline
            pTL_Inner.Pixels(6, 7).Color = pColorFill
            pTL_Inner.Pixels(6, 8).Color = pColorFill
            'row 7
            pTL_Inner.Pixels(7, 5).Color = pColorOutline
            pTL_Inner.Pixels(7, 6).Color = pColorFill
            pTL_Inner.Pixels(7, 7).Color = pColorFill
            pTL_Inner.Pixels(7, 8).Color = pColorFill
            'row 8
            pTL_Inner.Pixels(8, 5).Color = pColorOutline
            pTL_Inner.Pixels(8, 6).Color = pColorFill
            pTL_Inner.Pixels(8, 7).Color = pColorFill
            pTL_Inner.Pixels(8, 8).Color = pColorFill
        'Top-Right
            For iy = 1 To 8
                For ix = 1 To 8
                    pTR_Inner.Pixels(iy, ix).Color = pTL_Inner.Pixels(iy, (9 - ix)).Color
                Next ix
            Next iy
        'Bottom-Left
            For iy = 1 To 8
                For ix = 1 To 8
                    pBL_Inner.Pixels(iy, ix).Color = pTL_Inner.Pixels((9 - iy), ix).Color
                Next ix
            Next iy
        'Bottom-Right
            For iy = 1 To 8
                For ix = 1 To 8
                    pBR_Inner.Pixels(iy, ix).Color = pTL_Inner.Pixels((9 - iy), (9 - ix)).Color
                Next ix
            Next iy
    
    'STRAIGHT TILES
        'Left
            For iy = 1 To 8
                pL_Inner.Pixels(iy, 5).Color = pColorOutline
                pL_Inner.Pixels(iy, 6).Color = pColorFill
                pL_Inner.Pixels(iy, 7).Color = pColorFill
                pL_Inner.Pixels(iy, 8).Color = pColorFill
            Next iy
        'Right
            For iy = 1 To 8
                For ix = 1 To 8
                    pR_Inner.Pixels(iy, ix).Color = pL_Inner.Pixels(iy, (9 - ix)).Color
                Next ix
            Next iy
        'Top
            For ix = 1 To 8
                pT_Inner.Pixels(5, ix).Color = pColorOutline
                pT_Inner.Pixels(6, ix).Color = pColorFill
                pT_Inner.Pixels(7, ix).Color = pColorFill
                pT_Inner.Pixels(8, ix).Color = pColorFill
            Next ix
        'Bottom
            For iy = 1 To 8
                For ix = 1 To 8
                    pB_Inner.Pixels(iy, ix).Color = pT_Inner.Pixels((9 - iy), ix).Color
                Next ix
            Next iy
    
    'FILLED TILE
        'Blank
        For iy = 1 To 8
            For ix = 1 To 8
                pF_Blank.Pixels(iy, ix).Color = RGB(0, 0, 0) 'background black
            Next ix
        Next iy
        'Filled
        For iy = 1 To 8
            For ix = 1 To 8
                pF_Inner.Pixels(iy, ix).Color = pColorFill
            Next ix
        Next iy
    
    'INVERTED CORNER TILES
        'Top-Left
            'row 1
            pTL_InnerInv.Pixels(1, 5).Color = pColorOutline
            pTL_InnerInv.Pixels(1, 6).Color = pColorFill
            pTL_InnerInv.Pixels(1, 7).Color = pColorFill
            pTL_InnerInv.Pixels(1, 8).Color = pColorFill
            'row 2
            pTL_InnerInv.Pixels(2, 5).Color = pColorOutline
            pTL_InnerInv.Pixels(2, 6).Color = pColorFill
            pTL_InnerInv.Pixels(2, 7).Color = pColorFill
            pTL_InnerInv.Pixels(2, 8).Color = pColorFill
            'row 3
            pTL_InnerInv.Pixels(3, 5).Color = pColorOutline
            pTL_InnerInv.Pixels(3, 6).Color = pColorFill
            pTL_InnerInv.Pixels(3, 7).Color = pColorFill
            pTL_InnerInv.Pixels(3, 8).Color = pColorFill
            'row 4
            pTL_InnerInv.Pixels(4, 4).Color = pColorOutline
            pTL_InnerInv.Pixels(4, 5).Color = pColorFill
            pTL_InnerInv.Pixels(4, 6).Color = pColorFill
            pTL_InnerInv.Pixels(4, 7).Color = pColorFill
            pTL_InnerInv.Pixels(4, 8).Color = pColorFill
            'row 5
            pTL_InnerInv.Pixels(5, 1).Color = pColorOutline
            pTL_InnerInv.Pixels(5, 2).Color = pColorOutline
            pTL_InnerInv.Pixels(5, 3).Color = pColorOutline
            pTL_InnerInv.Pixels(5, 4).Color = pColorFill
            pTL_InnerInv.Pixels(5, 5).Color = pColorFill
            pTL_InnerInv.Pixels(5, 6).Color = pColorFill
            pTL_InnerInv.Pixels(5, 7).Color = pColorFill
            pTL_InnerInv.Pixels(5, 8).Color = pColorFill
            'row 6
            For ix = 1 To 8
                pTL_InnerInv.Pixels(6, ix).Color = pColorFill
            Next ix
            'row 7
            For ix = 1 To 8
                pTL_InnerInv.Pixels(7, ix).Color = pColorFill
            Next ix
            'row 8
            For ix = 1 To 8
                pTL_InnerInv.Pixels(8, ix).Color = pColorFill
            Next ix
        'Top-Right
            For iy = 1 To 8
                For ix = 1 To 8
                    pTR_InnerInv.Pixels(iy, ix).Color = pTL_InnerInv.Pixels(iy, (9 - ix)).Color
                Next ix
            Next iy
        'Bottom-Left
            For iy = 1 To 8
                For ix = 1 To 8
                    pBL_InnerInv.Pixels(iy, ix).Color = pTL_InnerInv.Pixels((9 - iy), ix).Color
                Next ix
            Next iy
        'Bottom-Right
            For iy = 1 To 8
                For ix = 1 To 8
                    pBR_InnerInv.Pixels(iy, ix).Color = pTL_InnerInv.Pixels((9 - iy), (9 - ix)).Color
                Next ix
            Next iy
    
    'INVERTED CORNER TILES WITH OUTLINE
        'Top-Left
            For iy = 1 To 8
                For ix = 1 To 8
                    pTL_InnerInvOL.Pixels(iy, ix).Color = pTL_InnerInv.Pixels(iy, ix).Color
                Next ix
            Next iy
            'apply outline row
            For ix = 1 To 8
                pTL_InnerInvOL.Pixels(8, ix).Color = pColorOutline
            Next ix
        'Top-Right
            For iy = 1 To 8
                For ix = 1 To 8
                    pTR_InnerInvOL.Pixels(iy, ix).Color = pTR_InnerInv.Pixels(iy, ix).Color
                Next ix
            Next iy
            'apply outline row
            For ix = 1 To 8
                pTR_InnerInvOL.Pixels(8, ix).Color = pColorOutline
            Next ix
        'Bottom-Left
            For iy = 1 To 8
                For ix = 1 To 8
                    pBL_InnerInvOL.Pixels(iy, ix).Color = pBL_InnerInv.Pixels(iy, ix).Color
                Next ix
            Next iy
            'apply outline row
            For ix = 1 To 8
                pBL_InnerInvOL.Pixels(1, ix).Color = pColorOutline
            Next ix
        'Bottom-Right
            For iy = 1 To 8
                For ix = 1 To 8
                    pBR_InnerInvOL.Pixels(iy, ix).Color = pBR_InnerInv.Pixels(iy, ix).Color
                Next ix
            Next iy
            'apply outline row
            For ix = 1 To 8
                pBR_InnerInvOL.Pixels(1, ix).Color = pColorOutline
            Next ix

End Sub

Private Sub Initialize_GhostCageTiles()

    'CORNER TILES
        'Top-Left
            'row 5
            pTL_GhostCage.Pixels(5, 5).Color = pColorOutline
            pTL_GhostCage.Pixels(5, 6).Color = pColorOutline
            pTL_GhostCage.Pixels(5, 7).Color = pColorOutline
            pTL_GhostCage.Pixels(5, 8).Color = pColorOutline
            'row 6
            pTL_GhostCage.Pixels(6, 5).Color = pColorOutline
            pTL_GhostCage.Pixels(6, 6).Color = pColorFill
            pTL_GhostCage.Pixels(6, 7).Color = pColorFill
            pTL_GhostCage.Pixels(6, 8).Color = pColorFill
            'row 7
            pTL_GhostCage.Pixels(7, 5).Color = pColorOutline
            pTL_GhostCage.Pixels(7, 6).Color = pColorFill
            pTL_GhostCage.Pixels(7, 7).Color = pColorFill
            pTL_GhostCage.Pixels(7, 8).Color = pColorFill
            'row 8
            pTL_GhostCage.Pixels(8, 5).Color = pColorOutline
            pTL_GhostCage.Pixels(8, 6).Color = pColorFill
            pTL_GhostCage.Pixels(8, 7).Color = pColorFill
            pTL_GhostCage.Pixels(8, 8).Color = pColorOutline
        'Top-Right
            For iy = 1 To 8
                For ix = 1 To 8
                    pTR_GhostCage.Pixels(iy, ix).Color = pTL_GhostCage.Pixels(iy, (9 - ix)).Color
                Next ix
            Next iy
        'Bottom-Left
            For iy = 1 To 8
                For ix = 1 To 8
                    pBL_GhostCage.Pixels(iy, ix).Color = pTL_GhostCage.Pixels((9 - iy), ix).Color
                Next ix
            Next iy
        'Bottom-Right
            For iy = 1 To 8
                For ix = 1 To 8
                    pBR_GhostCage.Pixels(iy, ix).Color = pTL_GhostCage.Pixels((9 - iy), (9 - ix)).Color
                Next ix
            Next iy
    
    'STRAIGHT TILES (reused straight outter tiles, but: T<=>B and L<=>R)
        'Left
            For iy = 1 To 8
                For ix = 1 To 8
                    pL_GhostCage.Pixels(iy, ix).Color = pR_Outter.Pixels(iy, ix).Color
                Next ix
            Next iy
        'Right
            For iy = 1 To 8
                For ix = 1 To 8
                    pR_GhostCage.Pixels(iy, ix).Color = pL_Outter.Pixels(iy, ix).Color
                Next ix
            Next iy
        'Top
            For iy = 1 To 8
                For ix = 1 To 8
                    pT_GhostCage.Pixels(iy, ix).Color = pB_Outter.Pixels(iy, ix).Color
                Next ix
            Next iy
        'Bottom
            For iy = 1 To 8
                For ix = 1 To 8
                    pB_GhostCage.Pixels(iy, ix).Color = pT_Outter.Pixels(iy, ix).Color
                Next ix
            Next iy
    
    'GHOST GATE
        'Gate Left
            'row 5
            pGL_GhostCage.Pixels(5, 1).Color = pColorOutline
            'row 6
            pGL_GhostCage.Pixels(6, 1).Color = pColorOutline
            pGL_GhostCage.Pixels(6, 2).Color = pColorGate
            pGL_GhostCage.Pixels(6, 3).Color = pColorGate
            pGL_GhostCage.Pixels(6, 4).Color = pColorGate
            pGL_GhostCage.Pixels(6, 5).Color = pColorGate
            pGL_GhostCage.Pixels(6, 6).Color = pColorGate
            pGL_GhostCage.Pixels(6, 7).Color = pColorGate
            pGL_GhostCage.Pixels(6, 8).Color = pColorGate
            'row 7
            pGL_GhostCage.Pixels(7, 1).Color = pColorOutline
            pGL_GhostCage.Pixels(7, 2).Color = pColorGate
            pGL_GhostCage.Pixels(7, 3).Color = pColorGate
            pGL_GhostCage.Pixels(7, 4).Color = pColorGate
            pGL_GhostCage.Pixels(7, 5).Color = pColorGate
            pGL_GhostCage.Pixels(7, 6).Color = pColorGate
            pGL_GhostCage.Pixels(7, 7).Color = pColorGate
            pGL_GhostCage.Pixels(7, 8).Color = pColorGate
            'row 8
            pGL_GhostCage.Pixels(8, 1).Color = pColorOutline
        'Gate Right
            For iy = 1 To 8
                For ix = 1 To 8
                    pGR_GhostCage.Pixels(iy, ix).Color = pGL_GhostCage.Pixels(iy, (9 - ix)).Color
                Next ix
            Next iy

End Sub
