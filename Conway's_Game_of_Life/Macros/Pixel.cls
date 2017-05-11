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
Private pAlive As Boolean 'True: cell is black and alive ; False: cell is white and dead
Private pN1 As Long 'Neighbor 1 color
Private pN2 As Long 'Neighbor 2 color
Private pN3 As Long 'Neighbor 3 color   N1  N2  N3
Private pN4 As Long 'Neighbor 4 color
Private pN6 As Long 'Neighbor 6 color   N4  Px  N6
Private pN7 As Long 'Neighbor 7 color
Private pN8 As Long 'Neighbor 8 color   N7  N8  N9
Private pN9 As Long 'Neighbor 9 color


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
' Pixel Alive or Dead
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Property Get Alive() As Boolean
    Alive = pAlive
End Property
Public Property Let Alive(Value As Boolean)
    pAlive = Value
End Property


'''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
' Neighbor 1
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Property Get N1() As Long
    N1 = pN1
End Property
Public Property Let N1(Value As Long)
    pN1 = Value
End Property


'''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
' Neighbor 2
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Property Get N2() As Long
    N2 = pN2
End Property
Public Property Let N2(Value As Long)
    pN2 = Value
End Property


'''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
' Neighbor 3
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Property Get N3() As Long
    N3 = pN3
End Property
Public Property Let N3(Value As Long)
    pN3 = Value
End Property


'''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
' Neighbor 4
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Property Get N4() As Long
    N4 = pN4
End Property
Public Property Let N4(Value As Long)
    pN4 = Value
End Property


'''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
' Neighbor 6
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Property Get N6() As Long
    N6 = pN6
End Property
Public Property Let N6(Value As Long)
    pN6 = Value
End Property


'''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
' Neighbor 7
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Property Get N7() As Long
    N7 = pN7
End Property
Public Property Let N7(Value As Long)
    pN7 = Value
End Property


'''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
' Neighbor 8
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Property Get N8() As Long
    N8 = pN8
End Property
Public Property Let N8(Value As Long)
    pN8 = Value
End Property


'''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
' Neighbor 9
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Property Get N9() As Long
    N9 = pN9
End Property
Public Property Let N9(Value As Long)
    pN9 = Value
End Property



'=================================================================================='
'==================================== Methods ====================================='
'=================================================================================='

Public Sub ReadNeighbors()

    pN1 = Wrks.Cells((pYrow - 1), (pXcol - 1)).Interior.Color
    pN2 = Wrks.Cells((pYrow - 1), (pXcol)).Interior.Color     ''''''''''''''''
    pN3 = Wrks.Cells((pYrow - 1), (pXcol + 1)).Interior.Color '  N1  N2  N3  '
    pN4 = Wrks.Cells((pYrow), (pXcol - 1)).Interior.Color     '              '
    'Px = Wrks.Cells((pYrow), (pXcol)).Inerior.Color          '  N4  Px  N6  '
    pN6 = Wrks.Cells((pYrow), (pXcol + 1)).Interior.Color     '              '
    pN7 = Wrks.Cells((pYrow + 1), (pXcol - 1)).Interior.Color '  N7  N8  N9  '
    pN8 = Wrks.Cells((pYrow + 1), (pXcol)).Interior.Color     ''''''''''''''''
    pN9 = Wrks.Cells((pYrow + 1), (pXcol + 1)).Interior.Color

End Sub

Public Function Items(ByVal index As Integer) As Long

    Dim rtrnColor As Long 'variable to hold color long to return
    
    'grab item's color
    If index = 1 Then
        rtrnColor = pN1
    ElseIf index = 2 Then
        rtrnColor = pN2
    ElseIf index = 3 Then
        rtrnColor = pN3
    ElseIf index = 4 Then
        rtrnColor = pN4
    ElseIf index = 5 Then
        rtrnColor = White
    ElseIf index = 6 Then
        rtrnColor = pN6
    ElseIf index = 7 Then
        rtrnColor = pN7
    ElseIf index = 8 Then
        rtrnColor = pN8
    ElseIf index = 9 Then
        rtrnColor = pN9
    End If
    
    'return item's color
    Items = rtrnColor

End Function

Public Sub Evolve()

    Dim i As Integer 'generic integer counter
    Dim nLiving As Integer 'number of living neighbors
    
    'initialize nLiving
    nLiving = 0
    
    'count number of living neighbors
    For i = 1 To 9
        If Me.Items(i) = Black Then: nLiving = nLiving + 1
    Next i
    
    'Living Pixel Rules
    If pAlive Then
        'Rule 1
        If nLiving < 2 Then
            pAlive = False
        'Rule 2
        ElseIf (nLiving = 2) Or (nLiving = 3) Then
            pAlive = True
        'Rule 3
        ElseIf nLiving > 3 Then
            pAlive = False
        End If
    'Dead Pixel Rules
    ElseIf Not pAlive Then
        'Rule 4
        If nLiving = 3 Then
            pAlive = True
        End If
    End If

End Sub
