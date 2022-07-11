VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TestMachine"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Implements Machine

Dim BaseMach As Machine
Public counter As Integer

Private Sub Class_Initialize()
    Set BaseMach = New Machine
End Sub

Private Sub Machine_Change(nState As State, Optional Mach As Machine = Nothing)
    BaseMach.Change nState, Me
End Sub

Private Sub Machine_Init(nState As State, Optional Mach As Machine = Nothing)
    BaseMach.Init nState, Me
End Sub

Private Sub Machine_Update(Optional Mach As Machine = Nothing)
    BaseMach.Update Me
End Sub

Sub Change(nState As State, Optional Mach As Machine = Nothing)
    Machine_Change nState, Me
End Sub
