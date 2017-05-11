VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Pixel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'=================================================================================='
'=================================== Properties ==================================='
'=================================================================================='

Private pXcol As Integer 'x or column coordinate
Private pYrow As Integer 'y or row coordinate
Private pColor As Long 'Color value of the pixel

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
' Color of Pixel
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Property Get Color() As Long
    Color = pColor
End Property
Public Property Let Color(Value As Long)
    pColor = Value
End Property


'=================================================================================='
'================================ Public Methods =================================='
'=================================================================================='

Public Sub Draw()
'adds pixel to its current color's corresponding range
    
    Dim i As Integer 'integer iterator
    
    'ignore black pixels since background is already black
    If pColor <> Black Then
        'find correct array to add to
        For i = 1 To UBound(CLR_Array)
        
            'add to range of this pixel's color
            If pColor = CR_Array((i + 1), 1) Then
                'add to range
                Set CLR_Array(i, 2) = Union(CLR_Array(i, 2), Cells(pYrow, pXcol))
                'increment counter
                CLR_Array(i, 1) = CLR_Array(i, 1) + 1
                'at 100 cells, dump and clear
                If CLR_Array(i, 1) = 100 Then
                    'dump
                    Set CR_Array((i + 1), 2) = Union(CR_Array((i + 1), 2), CLR_Array(i, 2))
                    'clear
                    Set CLR_Array(i, 2) = Cells(2, (i + 7))
                    CLR_Array(i, 1) = 0
                End If
            End If
            
        Next i
    End If

End Sub

Public Sub Map_Draw()
'adds pixel to its current color's corresponding MAP RANGE
    
    Dim i As Integer 'integer iterator
    
    'black pixel
    If pColor = Black Then
        'add to range
        Set MLR_Array(7, 2) = Union(MLR_Array(7, 2), Cells(pYrow, pXcol))
        'increment counter
        MLR_Array(7, 1) = MLR_Array(7, 1) + 1
        'at 100 cells, dump and clear
        If MLR_Array(7, 1) = 100 Then
            'dump
            Set CR_Array(1, 2) = Union(CR_Array(1, 2), MLR_Array(7, 2))
            'clear
            Set MLR_Array(7, 2) = Cells(2, 7)
            MLR_Array(7, 1) = 0
        End If
    End If
    
    'find correct map-range array to add to
    For i = 1 To UBound(MR_Array)
    
        'add to range of this pixel's color
        If pColor = MR_Array(i, 1) Then
            'add to range
            Set MLR_Array(i, 2) = Union(MLR_Array(i, 2), Cells(pYrow, pXcol))
            'increment counter
            MLR_Array(i, 1) = MLR_Array(i, 1) + 1
            'at 100 cells, dump and clear
            If MLR_Array(i, 1) = 100 Then
                'dump
                Set MR_Array(i, 2) = Union(MR_Array(i, 2), MLR_Array(i, 2))
                'clear
                Set MLR_Array(i, 2) = Cells(2, i)
                MLR_Array(i, 1) = 0
            End If
        End If
        
    Next i

End Sub
