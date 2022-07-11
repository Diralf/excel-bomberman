VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "GBomb"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Dim x As Integer
Dim y As Integer
Dim lifeTime As Integer
Dim currentLifeTime As Integer
Dim rangeFire As Integer
Dim directions(4, 2) As Integer
Dim FieldStart As String
Public ScoreCell As String

Private Sub Class_Initialize()
    directions(0, 0) = 1
    directions(0, 1) = 0
    
    directions(1, 0) = 0
    directions(1, 1) = -1
    
    directions(2, 0) = -1
    directions(2, 1) = 0
    
    directions(3, 0) = 0
    directions(3, 1) = 1
    
    lifeTime = 0
    currentLifeTime = 1
End Sub

Sub Init(ByVal nbomber As GBomber)
    x = nbomber.x
    y = nbomber.y
    lifeTime = nbomber.bombLifeTime
    currentLifeTime = 0
    rangeFire = nbomber.bombRange
    FieldStart = "B2"
    
    CellLib.Draw FieldStart, x, y, GameObjectsCode.EBOMB, RGB(0, 0, 0), RGB(255, 255, 255)
End Sub

Sub Bang()
    Dim TargetCell As String
    Dim targetX As Integer
    Dim targetY As Integer

    For I = 0 To 3
        For J = 0 To rangeFire
            targetX = x + J * directions(I, 0)
            targetY = y + J * directions(I, 1)
            If CellLib.IsOXYValid(targetX, targetY) Then
                Dim ValFromCell
                ValFromCell = CellLib.GetValFromCell(FieldStart, targetX, targetY)
                
                If ValFromCell = "1" Then Exit For
                
                CellLib.Draw FieldStart, targetX, targetY, "+", RGB(255, 100, 0), RGB(255, 255, 0)
                
                Select Case ValFromCell
                Case "2"
                    ScoreUp 1
                    Dim randRes As Integer
                    randRes = Int(10 * Rnd)
                    If randRes > 8 Then
                        CellLib.Draw FieldStart, targetX, targetY, "B", RGB(0, 180, 0), RGB(0, 0, 0)
                    ElseIf randRes > 6 Then
                        CellLib.Draw FieldStart, targetX, targetY, "F", RGB(0, 180, 0), RGB(0, 0, 0)
                    End If
                    Exit For
                Case "$"
                    ScoreUp 50
                Case "@"
                    ScoreUp 100
                End Select
                
            End If
        Next J
    Next I
End Sub

Sub ClearBang()
    Dim TargetCell As String
    Dim targetX As Integer
    Dim targetY As Integer

    For I = 0 To 3
        For J = 0 To rangeFire
            targetX = x + J * directions(I, 0)
            targetY = y + J * directions(I, 1)
            If CellLib.IsOXYValid(targetX, targetY) And CellLib.GetValFromCell(FieldStart, targetX, targetY) = "+" Then
                CellLib.Draw FieldStart, targetX, targetY, " ", RGB(255, 255, 255), RGB(255, 0, 0)
            End If
        Next J
    Next I
End Sub

Sub Update()
    If currentLifeTime <= lifeTime Then
        currentLifeTime = currentLifeTime + 1
        If currentLifeTime = lifeTime Then
            Bang
        End If
        
        If currentLifeTime < lifeTime And CellLib.GetValFromCell(FieldStart, x, y) = "+" Then
            Bang
            currentLifeTime = lifeTime
        End If
        
        If currentLifeTime > lifeTime Then
            ClearBang
        End If
    End If
End Sub

Public Function IsAlive() As Boolean
    IsAlive = currentLifeTime <= lifeTime
End Function

Sub ScoreUp(Score As Integer)
    CellLib.SetValToCell ScoreCell, 0, 0, CellLib.GetValFromCell(ScoreCell, 0, 0) + Score
End Sub
