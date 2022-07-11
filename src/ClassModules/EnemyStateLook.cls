VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "EnemyStateLook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Implements EnemyState

Private Sub EnemyState_InState(m As EnemyMachine)

End Sub

Private Sub EnemyState_OutState(m As EnemyMachine)

End Sub

Private Sub EnemyState_Update(m As EnemyMachine)
    If Not m.manager.IsAlive Then Exit Sub

    If CellLib.GetValFromCell(m.manager.FieldStart, m.manager.x, m.manager.y) = "+" Then
        m.manager.Die
        Exit Sub
    End If

    'If Int(5 * Rnd) > 3 Then
        m.manager.direction = m.manager.ChooseWay(Int(4 * Rnd))
    'End If
    m.manager.Move m.manager.direction
End Sub
