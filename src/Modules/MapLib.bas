Attribute VB_Name = "MapLib"


Function HandleCell(mapRange As range, fieldRange As range, J As Integer, I As Integer)
    Dim CellValue, TargetCell, TargetRange As range
    CellValue = mapRange.Cells(I + 1, J + 1).value
    
    Set TargetCell = fieldRange.Cells(I + 1, J + 1)
    Set TargetRange = range(TargetCell.address)
    
    
    
    Select Case CellValue
    Case 1
        CellLib.SetBackColorRange TargetRange, RGB(100, 40, 0)
        CellLib.SetFontColorRange TargetRange, RGB(100, 40, 0)
        TargetCell.value = CellValue
    Case 2
        Dim randRes As Integer
        randRes = Int(12 * Rnd)
        If randRes > 10 Then
            Game.enemies(Game.CountEnemy).Init New EnemyStateLook, J, I
            Game.CountEnemy = Game.CountEnemy + 1
        ElseIf randRes > 4 Then
            'CellLib.SetBackColorRange TargetRange, RGB(255, 170, 0)
            CellLib.SetBackColorRange TargetRange, RGB(170, 170, 170)
            CellLib.SetFontColorRange TargetRange, RGB(170, 170, 170)
            TargetCell.value = CellValue
        Else
            CellLib.SetBackColorRange TargetRange, RGB(255, 255, 255)
            CellLib.SetFontColorRange TargetRange, RGB(0, 0, 0)
            TargetCell.value = " "
        End If
    Case Else
        CellLib.SetBackColorRange TargetRange, RGB(255, 255, 255)
        CellLib.SetFontColorRange TargetRange, RGB(0, 0, 0)
        TargetCell.value = CellValue
    End Select
End Function

Function InitMap(MapStart As String, FieldStart As String, width As Integer, height As Integer)
    Dim I As Integer
    Dim J As Integer
    Dim mapRange As range
    Dim fieldRange As range
    Dim fieldEnd As String
    
    CountEnemy = 0
    
    Set mapRange = range(MapStart)
    Set fieldRange = range(FieldStart)
    
    fieldEnd = CellLib.ConvOXYToCell(FieldStart, width - 1, height - 1)
    range(FieldStart, fieldEnd).ClearContents
    
    
    For I = 0 To height - 1
        For J = 0 To width - 1
            HandleCell mapRange, fieldRange, J, I
        Next J
    Next I
End Function
