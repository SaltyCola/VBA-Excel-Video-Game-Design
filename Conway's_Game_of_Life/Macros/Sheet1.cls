VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    
    'only access sub if game not in progress
    If Not GameInProgress And PixelEditing Then
    
        Dim paintingWhite As Long 'white color variable for painting functionality
        Dim paintingBlack As Long 'black color variable for painting functionality
        
        'initialize color variables
        paintingWhite = RGB(255, 255, 255)
        paintingBlack = RGB(0, 0, 0)
        
        'change target's color to the opposite of what it currently is
        If Target.Interior.Color = paintingWhite Then
            Target.Interior.Color = paintingBlack
        ElseIf Target.Interior.Color = paintingBlack Then
            Target.Interior.Color = xlNone
        End If
    
    End If

End Sub
