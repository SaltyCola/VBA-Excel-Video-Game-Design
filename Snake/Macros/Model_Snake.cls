VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Model_Snake"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'=================================================================================='
'=================================== Properties ==================================='
'=================================================================================='

Private pLength As Integer 'determines how many cubes behind the head the tail should be
Private pHead As Snake_Head 'snake head object
Private pTail As Snake_Tail 'snake tail object
Private pEraser As Snake_Eraser 'snake eraser object
Private pBodySections As Collection 'collection of body section objects (# of entries is determined by length)

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
' Length of Snake
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Property Get Length() As Integer
    Length = pLength
End Property
Public Property Let Length(Value As Integer)
    pLength = Value
End Property

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
' Snake Head
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Property Get Head() As Snake_Head
    Set Head = pHead
End Property
Public Property Let Head(Value As Snake_Head)
    Set pHead = Value
End Property

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
' Snake Tail
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Property Get Tail() As Snake_Tail
    Set Tail = pTail
End Property
Public Property Let Tail(Value As Snake_Tail)
    Set pTail = Value
End Property

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
' Snake Eraser
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Property Get Eraser() As Snake_Eraser
    Set Eraser = pEraser
End Property
Public Property Let Eraser(Value As Snake_Eraser)
    Set pEraser = Value
End Property

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
' Snake Body Sections Collection
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Property Get BodySections() As Collection
    Set BodySections = pBodySections
End Property
Public Property Let BodySections(Value As Collection)
    Set pBodySections = Value
End Property


'================================================================================'
'==================================== Methods ==================================='
'================================================================================'

Private Sub Class_Initialize()
    
    'initialize subclasses
    Set pHead = New Snake_Head
    Set pTail = New Snake_Tail
    Set pEraser = New Snake_Eraser
    Set pBodySections = New Collection

End Sub

Public Sub Draw()

    Dim b As Integer 'iteration integer for drawing body sections
    
    
    'Draw new head position
    If DirUp Then 'Up
        Me.Head.CubeUp.DrawAt (yMin + Me.Head.Ypos), (xMin + Me.Head.Xpos)
    ElseIf DirLeft Then 'Left
        Me.Head.CubeLeft.DrawAt (yMin + Me.Head.Ypos), (xMin + Me.Head.Xpos)
    ElseIf DirRight Then 'Right
        Me.Head.CubeRight.DrawAt (yMin + Me.Head.Ypos), (xMin + Me.Head.Xpos)
    ElseIf DirDown Then 'Down
        Me.Head.CubeDown.DrawAt (yMin + Me.Head.Ypos), (xMin + Me.Head.Xpos)
    End If
    
    
    'draw body sections
    For b = 1 To Me.Length
        Me.BodySections.Item(b).CubeAny.DrawAt (yMin + Me.BodySections.Item(b).Ypos), (xMin + Me.BodySections.Item(b).Xpos)
    Next b
    
    
    'draw tail
    'up
    If (Me.Tail.Ypos = Me.BodySections.Item(Me.Length).Ypos + 3) And (Me.Tail.Xpos = Me.BodySections.Item(Me.Length).Xpos + 0) Then
        Me.Tail.CubeUp.DrawAt (yMin + Me.Tail.Ypos), (xMin + Me.Tail.Xpos)
    'left
    ElseIf (Me.Tail.Ypos = Me.BodySections.Item(Me.Length).Ypos + 0) And (Me.Tail.Xpos = Me.BodySections.Item(Me.Length).Xpos + 3) Then
        Me.Tail.CubeLeft.DrawAt (yMin + Me.Tail.Ypos), (xMin + Me.Tail.Xpos)
    'right
    ElseIf (Me.Tail.Ypos = Me.BodySections.Item(Me.Length).Ypos + 0) And (Me.Tail.Xpos = Me.BodySections.Item(Me.Length).Xpos - 3) Then
        Me.Tail.CubeRight.DrawAt (yMin + Me.Tail.Ypos), (xMin + Me.Tail.Xpos)
    'down
    ElseIf (Me.Tail.Ypos = Me.BodySections.Item(Me.Length).Ypos - 3) And (Me.Tail.Xpos = Me.BodySections.Item(Me.Length).Xpos + 0) Then
        Me.Tail.CubeDown.DrawAt (yMin + Me.Tail.Ypos), (xMin + Me.Tail.Xpos)
    End If
    
    'draw eraser
    Me.Eraser.CubeAny.DrawAt (yMin + Me.Eraser.Ypos), (xMin + Me.Eraser.Xpos)

End Sub

Public Sub CollisionAnimation()

    Dim b As Integer 'iteration integer for drawing body sections
    Dim stopX As Long 'x coordinate for stopping snake update
    Dim stopY As Long 'y coordinate for stopping snake update
    
    'set Snake velocity to zero
    Me.Head.Xvel = 0
    Me.Head.Yvel = 0
    
    'flash red
    Me.Head.RedFlash
    For b = 1 To Me.Length
        Me.BodySections.Item(b).RedFlash
    Next b
    Me.Tail.RedFlash
    
    'reverse updated positions
        'head
        stopX = Me.BodySections.Item(1).Xpos
        stopY = Me.BodySections.Item(1).Ypos
        Me.Head.Xpos = stopX
        Me.Head.Ypos = stopY
        'body sections
        For b = 1 To (Me.Length - 1)
            stopX = Me.BodySections.Item(b + 1).Xpos
            stopY = Me.BodySections.Item(b + 1).Ypos
            Me.BodySections.Item(b).Xpos = stopX
            Me.BodySections.Item(b).Ypos = stopY
        Next b
        stopX = Me.Tail.Xpos
        stopY = Me.Tail.Ypos
        Me.BodySections(Me.Length).Xpos = stopX
        Me.BodySections(Me.Length).Ypos = stopY
        'tail
        stopX = Me.Eraser.Xpos
        stopY = Me.Eraser.Ypos
        Me.Tail.Xpos = stopX
        Me.Tail.Ypos = stopY
        'eraser
        Me.Eraser.Xpos = 0
        Me.Eraser.Ypos = 0
        'draw snake
        Me.Draw
        
    'Game Over Message
    ActiveSheet.lblGameOver.Visible = True
    
    'hide control board
    ControlBoard.Hide

End Sub

Public Sub AddBody()

    Dim transBody As Snake_Body
    
    'increment length
    Me.Length = Me.Length + 1
    
    'initialize transfer Snake_Body object
    Set transBody = New Snake_Body
    
    'Load Cube
    transBody.SetCubes
    
    'set position
    If Me.Length > 1 Then
        'new body segment
        transBody.Xpos = Me.Tail.Xpos
        transBody.Ypos = Me.Tail.Ypos
    End If
    
    'set new keyBody name
    keyBody = "B" & Str(Me.Length)
    
    'add to BodySections
    Me.BodySections.Add transBody, keyBody
    
    'reset transBody
    Set transBody = Nothing

End Sub
