VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "GEnemy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public x As Integer
Public y As Integer
Public FieldStart As String
Public prevVal As String
Public IsAlive As Boolean
Public direction As Integer

Private Sub Class_Initialize()
    prevVal = " "
    IsAlive = False
End Sub

Sub Init(nx As Integer, ny As Integer)
    x = nx
    y = ny
    FieldStart = "B2"
    IsAlive = True
    direction = ChooseWay(Int(4 * Rnd))
    Draw
End Sub

Sub Draw()
    CellLib.Draw FieldStart, x, y, "$", RGB(255, 255, 255), RGB(255, 0, 255)
End Sub

Sub Die()
    CellLib.Draw FieldStart, x, y, ".", RGB(255, 255, 255), RGB(255, 0, 255)
    IsAlive = False
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
    Case "+"
        Die
    Case GameObjectsCode.EWALL, GameObjectsCode.EROCK, GameObjectsCode.EBOMB, "$"
        x = prevX
        y = prevY
        direction = ChooseWay(Int(4 * Rnd))
    Case Else
        prevVal = CellLib.GetCell(FieldStart, x, y).value
    End Select
End Sub

Sub Update()
    If Not IsAlive Then Exit Sub

    If CellLib.GetValFromCell(FieldStart, x, y) = "+" Then
        Die
        Exit Sub
    End If

    'If Int(5 * Rnd) > 3 Then
        direction = ChooseWay(Int(4 * Rnd))
    'End If
    Move direction
    Draw
End Sub

Function ChooseWay(direction As Integer) As Integer
    Dim curDir As Integer, xr As Integer, yr As Integer
    
    For I = 0 To 3
        curDir = direction + I
        If curDir > 3 Then curDir = curDir - 4
        
        Select Case curDir
        Case 0
            xr = x + 1
        Case 1
            yr = y - 1
        Case 2
            xr = x - 1
        Case 3
            yr = y + 1
        End Select
        
        If Not IsCollision(FieldStart, xr, yr) Then
            ChooseWay = curDir
        End If
    Next I
End Function

Function IsCollision(FieldStart As String, x As Integer, y As Integer) As Boolean
    Select Case CellLib.GetCell(FieldStart, x, y).value
    Case GameObjectsCode.EWALL, GameObjectsCode.EROCK, GameObjectsCode.EBOMB, "$"
        IsCollision = True
    Case Else
        IsCollision = False
    End Select
End Function
