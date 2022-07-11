VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "State2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Implements State

Private Sub State_InState(m As Machine)
    MsgBox "State2: Hi. I'm State2"
End Sub

Private Sub State_Update(m As Machine)
    If m.counter = 1 Then
        MsgBox "State2: Wait!"
        m.counter = m.counter + 1
    Else
        MsgBox "State2: So... Now I Will Change self on State1"
        m.Change New State1
    End If
End Sub

Private Sub State_OutState(m As Machine)
    MsgBox "State2: Good luck"
End Sub
