VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Machine"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Dim currentState As State
Public counter As Integer

Sub Init(nState As State)
    counter = 0
    Set currentState = nState
    currentState.InState Me
End Sub

Sub Update()
    currentState.Update Me
End Sub

Sub Change(nState As State)
    currentState.OutState Me
    Set currentState = nState
    currentState.InState Me
End Sub
