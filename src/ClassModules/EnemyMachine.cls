VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "EnemyMachine"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Dim currentState As EnemyState
Public counter As Integer
Public manager As GEnemy

Sub Init(nState As EnemyState, nx As Integer, ny As Integer)
    counter = 0
    Set manager = New GEnemy
    manager.Init nx, ny
    Set currentState = nState
    currentState.InState Me
End Sub

Sub Update()
    currentState.Update Me
    manager.Update
End Sub

Sub Change(nState As EnemyState)
    currentState.OutState Me
    Set currentState = nState
    currentState.InState Me
End Sub
