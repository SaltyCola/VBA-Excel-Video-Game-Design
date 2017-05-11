VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DrawTool_Cube"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'=================================================================================='
'=================================== Properties ==================================='
'=================================================================================='

Private pCx As Double 'x coord for center of cube
Private pCy As Double 'y coord for center of cube
Private pC1 As Long 'color of C1
Private pC2 As Long 'color of C2
Private pC3 As Long 'color of C3
Private pC4 As Long 'color of C4
Private pC5 As Long 'color of C5
Private pC6 As Long 'color of C6
Private pC7 As Long 'color of C7
Private pC8 As Long 'color of C8
Private pC9 As Long 'color of C9

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
' X and Y coordinates
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Property Get Cx() As Double
    Cx = pCx
End Property
Public Property Let Cx(Value As Double)
    pCx = Value
End Property

Public Property Get Cy() As Double
    Cy = pCy
End Property
Public Property Let Cy(Value As Double)
    pCy = Value
End Property

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
' Cube Cells
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'''''''''''''
' 1 ' 2 ' 3 '
'''''''''''''
' 4 ' 5 ' 6 '
'''''''''''''
' 7 ' 8 ' 9 '
'''''''''''''

Public Property Get C1() As Long
    C1 = pC1
End Property
Public Property Let C1(Value As Long)
    pC1 = Value
End Property

Public Property Get C2() As Long
    C2 = pC2
End Property
Public Property Let C2(Value As Long)
    pC2 = Value
End Property

Public Property Get C3() As Long
    C3 = pC3
End Property
Public Property Let C3(Value As Long)
    pC3 = Value
End Property

Public Property Get C4() As Long
    C4 = pC4
End Property
Public Property Let C4(Value As Long)
    pC4 = Value
End Property

Public Property Get C5() As Long
    C5 = pC5
End Property
Public Property Let C5(Value As Long)
    pC5 = Value
End Property

Public Property Get C6() As Long
    C6 = pC6
End Property
Public Property Let C6(Value As Long)
    pC6 = Value
End Property

Public Property Get C7() As Long
    C7 = pC7
End Property
Public Property Let C7(Value As Long)
    pC7 = Value
End Property

Public Property Get C8() As Long
    C8 = pC8
End Property
Public Property Let C8(Value As Long)
    pC8 = Value
End Property

Public Property Get C9() As Long
    C9 = pC9
End Property
Public Property Let C9(Value As Long)
    pC9 = Value
End Property


'================================================================================'
'==================================== Methods ==================================='
'================================================================================'

Public Sub DrawAt(ByVal CenterRowNumber As Integer, ByVal CenterColumnNumber As Integer)

    'assign cube center coordinates
    Me.Cy = CenterRowNumber
    Me.Cx = CenterColumnNumber

    'draw cube colors
    Cells((Me.Cy - 1), (Me.Cx - 1)).Interior.Color = Me.C1
    Cells((Me.Cy - 1), (Me.Cx)).Interior.Color = Me.C2      '''''''''''''
    Cells((Me.Cy - 1), (Me.Cx + 1)).Interior.Color = Me.C3  ' 1 ' 2 ' 3 '
    Cells((Me.Cy), (Me.Cx - 1)).Interior.Color = Me.C4      '''''''''''''
    Cells((Me.Cy), (Me.Cx)).Interior.Color = Me.C5          ' 4 ' 5 ' 6 '
    Cells((Me.Cy), (Me.Cx + 1)).Interior.Color = Me.C6      '''''''''''''
    Cells((Me.Cy + 1), (Me.Cx - 1)).Interior.Color = Me.C7  ' 7 ' 8 ' 9 '
    Cells((Me.Cy + 1), (Me.Cx)).Interior.Color = Me.C8      '''''''''''''
    Cells((Me.Cy + 1), (Me.Cx + 1)).Interior.Color = Me.C9

End Sub
