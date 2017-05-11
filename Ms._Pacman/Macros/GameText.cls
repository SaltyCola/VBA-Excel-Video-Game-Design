VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "GameText"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'=================================================================================='
'=================================== Properties ==================================='
'=================================================================================='

'Coordinates
Private pXcol As Integer 'x, or column, coordinate of the top-left-most pixel
Private pYrow As Integer 'y, or row, coordinate of the top-left-most pixel

'Text Properties
Private pLeftJustified As Boolean 'True: Text will be left justified
Private pRightJustified As Boolean 'True: Text will be right justified
Private pTextColor As Long 'Color of this text object
Private pTextCount As Integer 'Number of characters in this text object
Private pTextString As String 'String given to the Text object to generate the Text Cube Array
Private pText() As Cube 'Array holding one cube object for every character

'Text Tiles
Private pTextTilesList(1 To 38, 1 To 2) As Variant 'List holding each of the text tiles with their seed string
Private pcA As Cube                                     'Private property to Class.
Private pcB As Cube
Private pcC As Cube
Private pcD As Cube
Private pcE As Cube
Private pcF As Cube
Private pcG As Cube
Private pcH As Cube
Private pcI As Cube
Private pcJ As Cube
Private pcK As Cube
Private pcL As Cube
Private pcM As Cube
Private pcN As Cube
Private pcO As Cube
Private pcP As Cube
Private pcQ As Cube
Private pcR As Cube
Private pcS As Cube
Private pcT As Cube
Private pcU As Cube
Private pcV As Cube
Private pcW As Cube
Private pcX As Cube
Private pcY As Cube
Private pcZ As Cube
Private pcSpace As Cube
Private pcExclmPnt As Cube
Private pc0 As Cube
Private pc1 As Cube
Private pc2 As Cube
Private pc3 As Cube
Private pc4 As Cube
Private pc5 As Cube
Private pc6 As Cube
Private pc7 As Cube
Private pc8 As Cube
Private pc9 As Cube

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
' Left Justified
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Property Get LeftJustified() As Boolean
    LeftJustified = pLeftJustified
End Property
Public Property Let LeftJustified(Value As Boolean)
    pLeftJustified = Value
End Property

'''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
' Right Justified
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Property Get RightJustified() As Boolean
    RightJustified = pRightJustified
End Property
Public Property Let RightJustified(Value As Boolean)
    pRightJustified = Value
End Property

'''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
' Text Color
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Property Get TextColor() As Long
    TextColor = pTextColor
End Property
Public Property Let TextColor(Value As Long)
    pTextColor = Value
End Property

'''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
' Text Count
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Property Get TextCount() As Integer
    TextCount = pTextCount
End Property
Public Property Let TextCount(Value As Integer)
    pTextCount = Value
End Property

'''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
' Text String
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Property Get TextString() As String
    TextString = pTextString
End Property
Public Property Let TextString(Value As String)
    pTextString = Value
End Property

'''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
' Text Cubes Array
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Property Get Text(ByVal index As Integer) As Cube
    Set Text = pText(index)
End Property
Public Property Let Text(ByVal index As Integer, Value As Cube)
    Set pText(index) = Value
End Property


'================================================================================='
'================================ Public Methods ================================='
'================================================================================='

Public Sub GenerateText(ByVal txtString As String, ByVal iYrow As Integer, ByVal iXcol As Integer)

    Dim i As Integer 'iterator
    Dim ix As Integer 'iterator for pixel columns
    Dim iy As Integer 'iterator for pixel rows
    Dim tCube As Cube 'temporary cube object for copying tiles into text array
    
    'set text object coordinates
    pXcol = iXcol
    pYrow = iYrow
    
    'set string to object
    pTextString = txtString
    
    'get length
    pTextCount = Len(txtString)
    
    'resize cube array
    ReDim pText(1 To pTextCount) As Cube
    
    'generate text cube objects
    For i = 1 To pTextCount
        'grab tile
        Set tCube = Me.Tile(LCase(Mid(pTextString, i, 1)))
        'initialize text array cube
        Set pText(i) = New Cube
        'update position values
            'left justified
            If pLeftJustified Then
                pText(i).Yrow = pYrow
                pText(i).Xcol = pXcol + (8 * (i - 1))
                pText(i).SetPixels
            'right justified
            ElseIf pRightJustified Then
                pText(i).Yrow = pYrow
                pText(i).Xcol = pXcol - (8 * ((pTextCount - i + 1) - 1))
                pText(i).SetPixels
            End If
        'copy pixel colors
        For iy = 1 To 8
            For ix = 1 To 8
                pText(i).Pixels(iy, ix).Color = tCube.Pixels(iy, ix).Color
            Next ix
        Next iy
    Next i

End Sub

Public Sub InitializePixelColor(ByVal txtColor As Long)

    Dim i As Integer 'iterator
    Dim i1 As Integer 'iterator
    Dim ix As Integer 'iterator for pixel columns
    Dim iy As Integer 'iterator for pixel rows
    Dim tCube As Cube 'temporary cube for giving text color to pixels
    Dim pxRow As Integer 'Pixel Row to apply color
    Dim pXcol As Integer 'Pixel Column to apply color
    
    'set text color
    pTextColor = txtColor
    
    'iterate text tiles list
    For i = 1 To 38
        'grab current tile
        Set tCube = pTextTilesList(i, 1)
        'reset counters
        pxRow = 1
        pXcol = 1
        'iterate binary seed
        For i1 = 1 To 71
            'color pixel
            If Mid(pTextTilesList(i, 2), i1, 1) = "1" Then
                tCube.Pixels(pxRow, pXcol).Color = pTextColor
                pXcol = pXcol + 1
            'do not color pixel
            ElseIf Mid(pTextTilesList(i, 2), i1, 1) = "0" Then
                pXcol = pXcol + 1
            'move on to next row
            ElseIf Mid(pTextTilesList(i, 2), i1, 1) = ":" Then
                pxRow = pxRow + 1
                pXcol = 1
            End If
        Next i1
        'copy colors into text tiles list array
        For iy = 1 To 8
            For ix = 1 To 8
                pTextTilesList(i, 1).Pixels(iy, ix).Color = tCube.Pixels(iy, ix).Color
            Next ix
        Next iy
    Next i

End Sub

Public Sub DrawText()
'Draw to Color Range Arrays

    Dim i As Integer 'iterator
    
    'iterate text array
    For i = 1 To pTextCount
        pText(i).Draw
    Next i

End Sub

Public Sub Map_DrawText()
'Draw to Map Range Arrays

    Dim i As Integer 'iterator
    
    'iterate text array
    For i = 1 To pTextCount
        pText(i).Map_Draw
    Next i

End Sub

Public Sub SwitchJustification(ByVal boolLeft As Boolean, ByVal boolright As Boolean)

    'default to left justification in case of error
    If boolLeft And boolright Then
        boolLeft = True
        boolright = False
    ElseIf Not boolLeft And Not boolright Then
        boolLeft = True
        boolright = False
    End If
    
    'apply justification
    pLeftJustified = boolLeft
    pRightJustified = boolright

End Sub

Public Function Tile(ByVal char As String) As Cube

    'Grab Tile by code
    If char = "a" Then
        Set Tile = pcA
    ElseIf char = "b" Then
        Set Tile = pcB
    ElseIf char = "c" Then
        Set Tile = pcC
    ElseIf char = "d" Then
        Set Tile = pcD
    ElseIf char = "e" Then
        Set Tile = pcE
    ElseIf char = "f" Then
        Set Tile = pcF
    ElseIf char = "g" Then
        Set Tile = pcG
    ElseIf char = "h" Then
        Set Tile = pcH
    ElseIf char = "i" Then
        Set Tile = pcI
    ElseIf char = "j" Then
        Set Tile = pcJ
    ElseIf char = "k" Then
        Set Tile = pcK
    ElseIf char = "l" Then
        Set Tile = pcL
    ElseIf char = "m" Then
        Set Tile = pcM
    ElseIf char = "n" Then
        Set Tile = pcN
    ElseIf char = "o" Then
        Set Tile = pcO
    ElseIf char = "p" Then
        Set Tile = pcP
    ElseIf char = "q" Then
        Set Tile = pcQ
    ElseIf char = "r" Then
        Set Tile = pcR
    ElseIf char = "s" Then
        Set Tile = pcS
    ElseIf char = "t" Then
        Set Tile = pcT
    ElseIf char = "u" Then
        Set Tile = pcU
    ElseIf char = "v" Then
        Set Tile = pcV
    ElseIf char = "w" Then
        Set Tile = pcW
    ElseIf char = "x" Then
        Set Tile = pcX
    ElseIf char = "y" Then
        Set Tile = pcY
    ElseIf char = "z" Then
        Set Tile = pcZ
    ElseIf char = " " Then
        Set Tile = pcSpace
    ElseIf char = "!" Then
        Set Tile = pcExclmPnt
    ElseIf char = "0" Then
        Set Tile = pc0
    ElseIf char = "1" Then
        Set Tile = pc1
    ElseIf char = "2" Then
        Set Tile = pc2
    ElseIf char = "3" Then
        Set Tile = pc3
    ElseIf char = "4" Then
        Set Tile = pc4
    ElseIf char = "5" Then
        Set Tile = pc5
    ElseIf char = "6" Then
        Set Tile = pc6
    ElseIf char = "7" Then
        Set Tile = pc7
    ElseIf char = "8" Then
        Set Tile = pc8
    ElseIf char = "9" Then
        Set Tile = pc9
    Else
        Set Tile = pcSpace
    End If

End Function


'================================================================================='
'================================ Private Methods ================================'
'================================================================================='

Private Sub Class_Initialize()

    'default to left justified
    pLeftJustified = True
    pRightJustified = False
    
    'initialize cube objects
    Set pcA = New Cube
    Set pcB = New Cube
    Set pcC = New Cube
    Set pcD = New Cube
    Set pcE = New Cube
    Set pcF = New Cube
    Set pcG = New Cube
    Set pcH = New Cube
    Set pcI = New Cube
    Set pcJ = New Cube
    Set pcK = New Cube
    Set pcL = New Cube
    Set pcM = New Cube
    Set pcN = New Cube
    Set pcO = New Cube
    Set pcP = New Cube
    Set pcQ = New Cube
    Set pcR = New Cube
    Set pcS = New Cube
    Set pcT = New Cube
    Set pcU = New Cube
    Set pcV = New Cube
    Set pcW = New Cube
    Set pcX = New Cube
    Set pcY = New Cube
    Set pcZ = New Cube
    Set pcSpace = New Cube
    Set pcExclmPnt = New Cube
    Set pc0 = New Cube
    Set pc1 = New Cube
    Set pc2 = New Cube
    Set pc3 = New Cube
    Set pc4 = New Cube
    Set pc5 = New Cube
    Set pc6 = New Cube
    Set pc7 = New Cube
    Set pc8 = New Cube
    Set pc9 = New Cube
    
    'initialize pixel object properties
    pcA.Xcol = 1
    pcA.Yrow = 1
    pcA.SetPixels
    pcB.Xcol = 1
    pcB.Yrow = 1
    pcB.SetPixels
    pcC.Xcol = 1
    pcC.Yrow = 1
    pcC.SetPixels
    pcD.Xcol = 1
    pcD.Yrow = 1
    pcD.SetPixels
    pcE.Xcol = 1
    pcE.Yrow = 1
    pcE.SetPixels
    pcF.Xcol = 1
    pcF.Yrow = 1
    pcF.SetPixels
    pcG.Xcol = 1
    pcG.Yrow = 1
    pcG.SetPixels
    pcH.Xcol = 1
    pcH.Yrow = 1
    pcH.SetPixels
    pcI.Xcol = 1
    pcI.Yrow = 1
    pcI.SetPixels
    pcJ.Xcol = 1
    pcJ.Yrow = 1
    pcJ.SetPixels
    pcK.Xcol = 1
    pcK.Yrow = 1
    pcK.SetPixels
    pcL.Xcol = 1
    pcL.Yrow = 1
    pcL.SetPixels
    pcM.Xcol = 1
    pcM.Yrow = 1
    pcM.SetPixels
    pcN.Xcol = 1
    pcN.Yrow = 1
    pcN.SetPixels
    pcO.Xcol = 1
    pcO.Yrow = 1
    pcO.SetPixels
    pcP.Xcol = 1
    pcP.Yrow = 1
    pcP.SetPixels
    pcQ.Xcol = 1
    pcQ.Yrow = 1
    pcQ.SetPixels
    pcR.Xcol = 1
    pcR.Yrow = 1
    pcR.SetPixels
    pcS.Xcol = 1
    pcS.Yrow = 1
    pcS.SetPixels
    pcT.Xcol = 1
    pcT.Yrow = 1
    pcT.SetPixels
    pcU.Xcol = 1
    pcU.Yrow = 1
    pcU.SetPixels
    pcV.Xcol = 1
    pcV.Yrow = 1
    pcV.SetPixels
    pcW.Xcol = 1
    pcW.Yrow = 1
    pcW.SetPixels
    pcX.Xcol = 1
    pcX.Yrow = 1
    pcX.SetPixels
    pcY.Xcol = 1
    pcY.Yrow = 1
    pcY.SetPixels
    pcZ.Xcol = 1
    pcZ.Yrow = 1
    pcZ.SetPixels
    pcSpace.Xcol = 1
    pcSpace.Yrow = 1
    pcSpace.SetPixels
    pcExclmPnt.Xcol = 1
    pcExclmPnt.Yrow = 1
    pcExclmPnt.SetPixels
    pc0.Xcol = 1
    pc0.Yrow = 1
    pc0.SetPixels
    pc1.Xcol = 1
    pc1.Yrow = 1
    pc1.SetPixels
    pc2.Xcol = 1
    pc2.Yrow = 1
    pc2.SetPixels
    pc3.Xcol = 1
    pc3.Yrow = 1
    pc3.SetPixels
    pc4.Xcol = 1
    pc4.Yrow = 1
    pc4.SetPixels
    pc5.Xcol = 1
    pc5.Yrow = 1
    pc5.SetPixels
    pc6.Xcol = 1
    pc6.Yrow = 1
    pc6.SetPixels
    pc7.Xcol = 1
    pc7.Yrow = 1
    pc7.SetPixels
    pc8.Xcol = 1
    pc8.Yrow = 1
    pc8.SetPixels
    pc9.Xcol = 1
    pc9.Yrow = 1
    pc9.SetPixels
    
    'initialize text tiles list array
    Set pTextTilesList(1, 1) = pcA
    pTextTilesList(1, 2) = "00000000:00111000:01101100:11000110:11000110:11111110:11000110:11000110"
    Set pTextTilesList(2, 1) = pcB
    pTextTilesList(2, 2) = "00000000:11111100:11000010:11000010:11111100:11000010:11000010:11111100"
    Set pTextTilesList(3, 1) = pcC
    pTextTilesList(3, 2) = "00000000:00111100:01100110:11000000:11000000:11000000:01100110:00111100"
    Set pTextTilesList(4, 1) = pcD
    pTextTilesList(4, 2) = "00000000:11111000:11001100:11000110:11000110:11000110:11001100:11111000"
    Set pTextTilesList(5, 1) = pcE
    pTextTilesList(5, 2) = "00000000:01111110:01100000:01100000:01111100:01100000:01100000:01111110"
    Set pTextTilesList(6, 1) = pcF
    pTextTilesList(6, 2) = "00000000:01111110:01100000:01100000:01111100:01100000:01100000:01100000"
    Set pTextTilesList(7, 1) = pcG
    pTextTilesList(7, 2) = "00000000:00111110:01100000:11000000:11001110:11000110:01100110:00111110"
    Set pTextTilesList(8, 1) = pcH
    pTextTilesList(8, 2) = "00000000:11000110:11000110:11000110:11111110:11000110:11000110:11000110"
    Set pTextTilesList(9, 1) = pcI
    pTextTilesList(9, 2) = "00000000:01111110:00011000:00011000:00011000:00011000:00011000:01111110"
    Set pTextTilesList(10, 1) = pcJ
    pTextTilesList(10, 2) = "00000000:01111110:00001100:00001100:00001100:00001100:10001100:01111000"
    Set pTextTilesList(11, 1) = pcK
    pTextTilesList(11, 2) = "00000000:11000110:11001100:11011000:11110000:11011000:11001100:11000110"
    Set pTextTilesList(12, 1) = pcL
    pTextTilesList(12, 2) = "00000000:11000000:11000000:11000000:11000000:11000000:11000000:11111110"
    Set pTextTilesList(13, 1) = pcM
    pTextTilesList(13, 2) = "00000000:11000110:11101110:11111110:11111110:11010110:11000110:11000110"
    Set pTextTilesList(14, 1) = pcN
    pTextTilesList(14, 2) = "00000000:11000110:11100110:11110110:11111110:11011110:11001110:11000110"
    Set pTextTilesList(15, 1) = pcO
    pTextTilesList(15, 2) = "00000000:01111100:11000110:11000110:11000110:11000110:11000110:01111100"
    Set pTextTilesList(16, 1) = pcP
    pTextTilesList(16, 2) = "00000000:11111100:11000110:11000110:11000110:11111100:11000000:11000000"
    Set pTextTilesList(17, 1) = pcQ
    pTextTilesList(17, 2) = "00000000:01111100:11000110:11000110:11010110:11001110:11000100:01111010"
    Set pTextTilesList(18, 1) = pcR
    pTextTilesList(18, 2) = "00000000:11111000:11000110:11000110:11001110:11111000:11011100:11001110"
    Set pTextTilesList(19, 1) = pcS
    pTextTilesList(19, 2) = "00000000:01111000:11001100:11000000:01111100:00000110:11000110:01111100"
    Set pTextTilesList(20, 1) = pcT
    pTextTilesList(20, 2) = "00000000:01111110:00011000:00011000:00011000:00011000:00011000:00011000"
    Set pTextTilesList(21, 1) = pcU
    pTextTilesList(21, 2) = "00000000:11000110:11000110:11000110:11000110:11000110:11000110:01111100"
    Set pTextTilesList(22, 1) = pcV
    pTextTilesList(22, 2) = "00000000:11000110:11000110:11000110:11101110:01111100:00111000:00010000"
    Set pTextTilesList(23, 1) = pcW
    pTextTilesList(23, 2) = "00000000:11000110:11000110:11010110:11111110:11111110:11101110:11000110"
    Set pTextTilesList(24, 1) = pcX
    pTextTilesList(24, 2) = "00000000:11000110:11000110:01101100:00111000:01101100:11000110:11000110"
    Set pTextTilesList(25, 1) = pcY
    pTextTilesList(25, 2) = "00000000:01100110:01100110:01100110:00111100:00011000:00011000:00011000"
    Set pTextTilesList(26, 1) = pcZ
    pTextTilesList(26, 2) = "00000000:01111110:00000110:00001100:00011000:00110000:01100000:01111110"
    Set pTextTilesList(27, 1) = pcSpace
    pTextTilesList(27, 2) = "00000000:00000000:00000000:00000000:00000000:00000000:00000000:00000000"
    Set pTextTilesList(28, 1) = pcExclmPnt
    pTextTilesList(28, 2) = "00000000:00111000:00111000:01110000:01100000:01000000:00000000:10000000"
    Set pTextTilesList(29, 1) = pc0
    pTextTilesList(29, 2) = "00000000:00111000:01001100:11000110:11000110:11000110:01100100:00111000"
    Set pTextTilesList(30, 1) = pc1
    pTextTilesList(30, 2) = "00000000:00011000:00111000:00011000:00011000:00011000:00011000:01111110"
    Set pTextTilesList(31, 1) = pc2
    pTextTilesList(31, 2) = "00000000:01111100:11000110:00001110:00111100:01111000:11100000:11111110"
    Set pTextTilesList(32, 1) = pc3
    pTextTilesList(32, 2) = "00000000:01111110:00001100:00011000:00111100:00000110:11000110:01111100"
    Set pTextTilesList(33, 1) = pc4
    pTextTilesList(33, 2) = "00000000:00011100:00111100:01101100:11001100:11111110:00001100:00001100"
    Set pTextTilesList(34, 1) = pc5
    pTextTilesList(34, 2) = "00000000:11111100:11000000:11111100:00000110:00000110:11000110:01111100"
    Set pTextTilesList(35, 1) = pc6
    pTextTilesList(35, 2) = "00000000:00111100:01100000:11000000:11111100:11000110:11000110:01111100"
    Set pTextTilesList(36, 1) = pc7
    pTextTilesList(36, 2) = "00000000:11111110:11000110:00001100:00011000:00110000:00110000:00110000"
    Set pTextTilesList(37, 1) = pc8
    pTextTilesList(37, 2) = "00000000:00111100:01100010:01110010:00111100:01001110:01000110:00111100"
    Set pTextTilesList(38, 1) = pc9
    pTextTilesList(38, 2) = "00000000:01111100:11000110:11000110:01111110:00000110:00001100:01111000"

End Sub
