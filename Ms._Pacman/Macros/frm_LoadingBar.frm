VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frm_LoadingBar 
   Caption         =   "Please Wait"
   ClientHeight    =   720
   ClientLeft      =   48
   ClientTop       =   372
   ClientWidth     =   2772
   OleObjectBlob   =   "frm_LoadingBar.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frm_LoadingBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Public Sub UpdateLoadingBar(ByVal LoadingMessage As String, ByVal indexCurrent As Integer, ByVal indexTotal As Integer)

    Dim i As Integer 'generic integer variable for iteration
    Dim lInt As Integer 'Loading Bar integer that will keep track of which img should be visible
    
    'update message
    Me.Label1.Caption = LoadingMessage
    
    'calculate new LInt
    lInt = Int((indexCurrent / indexTotal) * 166)
    
    'update loading image
    For i = 0 To 166
        Me.Controls("Image" & i).Visible = False
    Next i
    Me.Controls("Image" & lInt).Visible = True
    
    'show updates and maintain progression of code
    Me.Show vbModeless
    DoEvents
    
    'close loading bar if end reached
    If indexCurrent = indexTotal Then
        Me.Hide
    End If

End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
'Prevent userform close on red x click.

    If CloseMode = 0 Then: Cancel = True

End Sub
