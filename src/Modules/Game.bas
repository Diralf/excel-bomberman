Attribute VB_Name = "Game"
Dim Flag
Dim KeyVal As Integer
Dim Player2Val As Integer
Dim Cll
Dim pos As Integer

Dim posX As Integer
Dim posY As Integer

Dim StatusCell As String
Dim GameField As String
Dim GameFieldX As String
Dim GameFieldY As Integer

Dim bMan As GBomber
Dim bMan2 As GBomber
Dim bEnemy As GEnemy
Public enemies(200) As EnemyMachine
Public CountEnemy As Integer
Dim prevVal As String

Dim bombArr(2) As GBomb
Dim bomb2 As GBomb

Dim bomb3 As GBomb

Dim prevTime

Dim Mach As Machine

Function GlobalInit()
    StatusCell = "B1"
    GameField = "B2"
    GameFieldX = "B"
    GameFieldY = 2
    pos = 0
    Flag = True
    KeyVal = 99
    prevVal = " "

    posX = 5
    posY = 5

    For I = 0 To UBound(enemies())
        Set enemies(I) = New EnemyMachine
    Next I
End Function

Sub RunCycle()
Attribute RunCycle.VB_ProcData.VB_Invoke_Func = "t\n14"
    GLoop
End Sub

Sub GStop()
Attribute GStop.VB_ProcData.VB_Invoke_Func = "q\n14"
'
' MoveUp
'
' : Ctrl+q
'
    If StatusCell = Empty Then StatusCell = "B1"
    Flag = False
    range(StatusCell).FormulaR1C1 = "Stop"

    CellLib.SetBackColorCell StatusCell, 0, 0, RGB(250, 50, 50)
End Sub

Sub MoveUp()
Attribute MoveUp.VB_ProcData.VB_Invoke_Func = "w\n14"
    KeyVal = 1
    Update
End Sub
Sub MoveDown()
Attribute MoveDown.VB_ProcData.VB_Invoke_Func = "s\n14"
    KeyVal = 3
    Update
End Sub
Sub MoveLeft()
Attribute MoveLeft.VB_ProcData.VB_Invoke_Func = "a\n14"
    KeyVal = 2
    Update
End Sub
Sub MoveRight()
Attribute MoveRight.VB_ProcData.VB_Invoke_Func = "d\n14"
    KeyVal = 0
    Update
End Sub

Sub MoveUp2()
Attribute MoveUp2.VB_ProcData.VB_Invoke_Func = "i\n14"
    Player2Val = 1
    Update2
End Sub
Sub MoveDown2()
Attribute MoveDown2.VB_ProcData.VB_Invoke_Func = "k\n14"
    Player2Val = 3
    Update2
End Sub
Sub MoveLeft2()
Attribute MoveLeft2.VB_ProcData.VB_Invoke_Func = "j\n14"
    Player2Val = 2
    Update2
End Sub
Sub MoveRight2()
Attribute MoveRight2.VB_ProcData.VB_Invoke_Func = "l\n14"
    Player2Val = 0
    Update2
End Sub

Sub Attack()
Attribute Attack.VB_ProcData.VB_Invoke_Func = "f\n14"
    bMan.ThrowBomb
End Sub

Sub Attack2()
Attribute Attack2.VB_ProcData.VB_Invoke_Func = "p\n14"
    bMan2.ThrowBomb
End Sub

Function Update()
    bMan.Move (KeyVal)
    KeyVal = 99
End Function

Function Update2()
    bMan2.Move (Player2Val)
    Player2Val = 99
End Function

Function UpdateGameCycle()
    'Dim I As Integer
    'CellLib.SetValToCell "G2", pos, 0, "1"
    'pos = pos + 1
    'If pos > 9 Then pos = 0
    'CellLib.SetValToCell "G2", pos, 0, KeyVal
    'CellLib.SetBackColorCell "G2", pos, 0, RGB(100, 255, 255)


    For I = 0 To CountEnemy - 1
        enemies(I).Update
    Next I

    bMan.Update
    bMan2.Update

    'Mach.Update

End Function

Sub GLoop()
    If Not Flag Then Exit Sub
    UpdateGameCycle
    Application.OnTime Now + TimeValue("00:00:01"), "GLoop"
End Sub

Sub ResetGame()
    GStop
    Application.OnTime Now + TimeValue("00:00:01"), "GLoop", False
    Main
End Sub

Sub Main()
Attribute Main.VB_ProcData.VB_Invoke_Func = "r\n14"
    GlobalInit
    Set Mach = New Machine

    'Mach.Init New State1

    range(StatusCell).FormulaR1C1 = "Loading"
    CellLib.SetBackColorCell StatusCell, 0, 0, RGB(100, 100, 255)

    MapLib.InitMap "B32", "B2", 27, 27
    'MapLib.InitMap "B32", "B2", 63, 27

    Set bMan = New GBomber
    bMan.Init 1, 1, RGB(0, 0, 200), "M1", "R1"
    CellLib.SetValToCell "R1", 0, 0, 0

    Set bMan2 = New GBomber
    bMan2.Init 25, 25, RGB(200, 0, 0), "AB1", "AG1"
    CellLib.SetValToCell "AG1", 0, 0, 0

    range(StatusCell).FormulaR1C1 = "Running"
    CellLib.SetBackColorCell StatusCell, 0, 0, RGB(50, 200, 50)
    GLoop

End Sub
