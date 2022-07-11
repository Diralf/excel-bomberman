VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "State1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Implements State

Private Sub State_InState(m As Machine)
    MsgBox "State1: Hello! I'm State1"
End Sub

Private Sub State_Update(m As Machine)
    If m.counter = 0 Then
        MsgBox "State1: Now I will Change self to State 2"
        m.Change New State2
    Else
        MsgBox "State1: Now I Stay this till end"
    End If
End Sub

Private Sub State_OutState(m As Machine)
    MsgBox "State1: Bye-bye!"
    m.counter = m.counter + 1
End Sub
