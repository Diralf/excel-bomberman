VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "GBomber"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public x As Integer
Public y As Integer
Public countBomb As Integer
Public bombLifeTime As Integer
Public bombRange As Integer
Dim bombBag(15) As GBomb
Dim FieldStart As String
Dim prevVal As String
Dim PlayerColor As Long
Dim PlayerName As String
Dim PlayerCellName As String
Dim ScoreCell As String
Dim spawnX As Integer
Dim spawnY As Integer

Private Sub Class_Initialize()
    prevVal = " "
    
    For I = 0 To UBound(bombBag())
        Set bombBag(I) = New GBomb
    Next I
End Sub

Sub Init(nx As Integer, ny As Integer, pColor As Long, CellWithName As String, ScrCell As String)
    x = nx
    y = ny
    spawnX = nx
    spawnY = ny
    countBomb = 1
    bombLifeTime = 3
    bombRange = 2
    FieldStart = "B2"
    PlayerColor = pColor
    PlayerCellName = CellWithName
    PlayerName = CellLib.GetValFromCell(CellWithName, 0, 0)
    ScoreCell = ScrCell
    
    For I = 0 To UBound(bombBag())
        bombBag(I).ScoreCell = ScrCell
    Next I
    
    Draw
End Sub

Sub Draw()
    CellLib.Draw FieldStart, x, y, "@", RGB(255, 255, 255), PlayerColor
End Sub

Sub Move(direction As Integer)
    Dim prevX As Integer
    Dim prevY As Integer
    prevX = x
    prevY = y
    
    CellLib.GetCell(FieldStart, x, y).value = prevVal
    
    Select Case direction
    Case 0
        x = x + 1
    Case 1
        y = y - 1
    Case 2
        x = x - 1
    Case 3
        y = y + 1
    End Select
    
    Select Case CellLib.GetCell(FieldStart, x, y).value
    Case "+", "$"
        Die
    Case GameObjectsCode.EWALL, GameObjectsCode.EROCK, GameObjectsCode.EBOMB
        x = prevX
        y = prevY
    Case "B"
        TakeBonus GameBonus.BBOMB
    Case "F"
        TakeBonus GameBonus.BFIRE
    Case Else
        prevVal = CellLib.GetCell(FieldStart, x, y).value
    End Select
    Draw
End Sub

Sub ThrowBomb()
    Dim ActiveBomb As GBomb
    Dim IsFound As Boolean
    IsFound = False
    
    For I = 0 To countBomb - 1
        If Not bombBag(I).IsAlive Then
            Set ActiveBomb = bombBag(I)
            IsFound = True
            Exit For
        End If
    Next I

    If IsFound Then
        ActiveBomb.Init Me
        prevVal = CellLib.GetCell(FieldStart, x, y).value
    End If
    
End Sub

Sub Die()
    CellLib.Draw FieldStart, x, y, "x", RGB(255, 0, 0), RGB(255, 255, 255)
    'MsgBox PlayerName & " Lose"
    'Game.ResetGame
    ScoreUp -200
    Init spawnX, spawnY, PlayerColor, PlayerCellName, ScoreCell
End Sub

Sub TakeBonus(Bonus As Long)
    ScoreUp 10
    Select Case Bonus
    Case GameBonus.BBOMB
        If countBomb < 15 Then countBomb = countBomb + 1
    Case GameBonus.BFIRE
        bombRange = bombRange + 1
    End Select
End Sub

Sub ScoreUp(Score As Integer)
    CellLib.SetValToCell ScoreCell, 0, 0, CellLib.GetValFromCell(ScoreCell, 0, 0) + Score
End Sub

Sub Update()
    For I = 0 To UBound(bombBag())
        bombBag(I).Update
    Next I
    
    Select Case CellLib.GetValFromCell(FieldStart, x, y)
    Case "+", "$"
        Die
    End Select
End Sub


