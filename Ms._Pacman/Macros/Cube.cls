VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Cube"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'=================================================================================='
'=================================== Properties ==================================='
'=================================================================================='

'[FOR PACDOT CUBES ONLY]===================================================
'Cube's Dot Type
Private pDotType As Integer '0 = None, 1 = PacDot, 2 = PowerDot
Private pDotColor As Long 'PacDot Color taken from map object
'[FOR PACDOT CUBES ONLY]===================================================

'Cube Coordinates (Top-Left Corner)
Private pXcol As Integer 'x, or column, coordinate of the top-left-most pixel in the cube
Private pYrow As Integer 'y, or row, coordinate of the top-left-most pixel in the cube

'Pixels Array
Private pPixels(1 To 8, 1 To 8) As Pixel 'contains all 64 pixel objects per cube

'''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
' Cube's Dot Type
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Property Get DotType() As Integer
    DotType = pDotType
End Property
Public Property Let DotType(Value As Integer)
    pDotType = Value
End Property

'''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
' Cube's PacDot Color
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Property Get DotColor() As Long
    DotColor = pDotColor
End Property
Public Property Let DotColor(Value As Long)
    pDotColor = Value
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
' Pixels Array
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Property Get Pixels(ByVal iRow As Integer, ByVal iCol As Integer) As Pixel
    Set Pixels = pPixels(iRow, iCol)
End Property
Public Property Let Pixels(ByVal iRow As Integer, ByVal iCol As Integer, Value As Pixel)
    Set pPixels(iRow, iCol) = Value
End Property


'=================================================================================='
'==================================== Methods ====================================='
'=================================================================================='

Private Sub Class_Initialize()

    'initialize dot type property
    pDotType = 0

End Sub

Public Sub SetPixels()
'only works if cube's coordinates are set first

    Dim ir As Integer 'integer iterator for rows of pixels array
    Dim ic As Integer 'integer iterator for columns of pixels array
    Dim tpX As Pixel 'temporary pixel object
    
    'iterate pixels array
    For ir = 1 To 8
        For ic = 1 To 8
            'initialize pixel object
            Set tpX = New Pixel
            'set pixel coordinates
            tpX.Yrow = Me.Yrow + (ir - 1)
            tpX.Xcol = Me.Xcol + (ic - 1)
            'add pixel object to pixels array
            Set pPixels(ir, ic) = tpX
        Next ic
    Next ir

End Sub

Public Sub UpdatePixels()
'updates pixel coords during sprite movement

    Dim ir As Integer 'integer iterator for rows of pixels array
    Dim ic As Integer 'integer iterator for columns of pixels array
    Dim tpX As Pixel 'temporary pixel object
    
    'iterate pixels array
    For ir = 1 To 8
        For ic = 1 To 8
            'initialize pixel object
            Set tpX = pPixels(ir, ic)
            'set pixel coordinates
            tpX.Yrow = Me.Yrow + (ir - 1)
            tpX.Xcol = Me.Xcol + (ic - 1)
            'add pixel object to pixels array
            Set pPixels(ir, ic) = tpX
        Next ic
    Next ir

End Sub

Public Sub Draw()

    Dim ir As Integer 'integer iterator for rows of pixels array
    Dim ic As Integer 'integer iterator for columns of pixels array
    Dim tpX As Pixel 'temporary pixel object
    
    'iterate pixels array
    For ir = 1 To 8
        For ic = 1 To 8
            'grab pixel object
            Set tpX = pPixels(ir, ic)
            'draw pixel
            tpX.Draw
        Next ic
    Next ir

End Sub

Public Sub Map_Draw()

    Dim ir As Integer 'integer iterator for rows of pixels array
    Dim ic As Integer 'integer iterator for columns of pixels array
    Dim tpX As Pixel 'temporary pixel object
    
    'iterate pixels array
    For ir = 1 To 8
        For ic = 1 To 8
            'grab pixel object
            Set tpX = pPixels(ir, ic)
            'if pacdot, recolor 4 pixels at pacdot pixel locations
            If ir = 4 Or ir = 5 Then
                If ic = 4 Or ic = 5 Then
                    If pDotType = 1 Then
                        tpX.Color = pDotColor
                    End If
                End If
            End If
            'draw pixel
            tpX.Map_Draw
        Next ic
    Next ir

End Sub
