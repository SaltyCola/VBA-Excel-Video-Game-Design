Attribute VB_Name = "Testing"
Sub ClassesTest()

    Dim Snakey As Model_Snake
    Dim transBody As Snake_Body
    
    Set Snakey = New Model_Snake
    Set transBody = New Snake_Body
    
    Snakey.Head.Xpos = 2
    Snakey.Head.Ypos = 2
    Snakey.Head.Xvel = 0
    Snakey.Head.Yvel = 1
    
    Snakey.Tail.Xpos = 3
    Snakey.Tail.Ypos = 4
    Snakey.Tail.Xvel = 0
    Snakey.Tail.Yvel = -1
    
    Snakey.Length = 1
    
    transBody.Xpos = 5
    transBody.Ypos = 9
    transBody.Xvel = -1
    transBody.Yvel = 0
    
    Snakey.AddBody transBody
'    snakey.BodySections.Count
    MsgBox Snakey.BodySections.Item(1).Xpos
'    snakey.BodySections.Remove ("B1")
    
    
End Sub
