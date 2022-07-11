Attribute VB_Name = "CellLib"
Function GetCell(StartCell As String, x As Integer, y As Integer)
    Set GetCell = range(StartCell).Cells(y + 1, x + 1)
End Function

Function IsOXYValid(x As Integer, y As Integer) As Boolean
    IsOXYValid = x >= 0 And y >= 0
End Function

' need
Function SetValToCell(StartCell As String, x As Integer, y As Integer, newValue)
    If Not IsOXYValid(x, y) Then Exit Function
    range(StartCell).Cells(y + 1, x + 1).value = newValue
End Function

' need
Function GetValFromCell(StartCell As String, x As Integer, y As Integer)
    If Not IsOXYValid(x, y) Then Exit Function
    GetValFromCell = range(StartCell).Cells(y + 1, x + 1).value
End Function

Function SetBackColorRange(TargetRange As range, RGBColor As Long)
    With TargetRange.Interior
        .Color = RGBColor
    End With
End Function

Function SetFontColorRange(TargetRange As range, RGBColor As Long)
    With TargetRange.Font
        .Color = RGBColor
    End With
End Function

'need
Function SetBackColorCell(StartCell As String, x As Integer, y As Integer, RGBColor As Long)
    With range(ConvOXYToCell(StartCell, x, y)).Interior
        .Color = RGBColor
    End With
End Function

'need
Function SetFontColorCell(StartCell As String, x As Integer, y As Integer, RGBColor As Long)
    With range(ConvOXYToCell(StartCell, x, y)).Font
        .Color = RGBColor
    End With
End Function

Function Draw(FieldStart As String, x As Integer, y As Integer, CellValue, BackColor As Long, FontColor As Long)
    SetValToCell FieldStart, x, y, CellValue
    SetFontColorCell FieldStart, x, y, FontColor
    SetBackColorCell FieldStart, x, y, BackColor
End Function

'?
Function GetCellOX(Cell As String) As Integer
    GetCellOX = range(Cell).Column
End Function

'?
Function GetCellOY(Cell As String) As Integer
    GetCellOY = range(Cell).Row
End Function

' ?
Function GetCellAddress(x As Integer, y As Integer) As String
    GetCellAddress = range("A1").Cells(y, x).address
End Function

'need
Function ConvOXYToCell(StartCell As String, x As Integer, y As Integer)
    If Not IsOXYValid(x, y) Then Exit Function
    ConvOXYToCell = range(StartCell).Cells(y + 1, x + 1).address
End Function

Function ConvCellToOX(StartCell As String, Cell As String)
    ConvCellToOX = GetCellOX(Cell) - GetCellOX(StartCell)
End Function

Function ConvCellToOY(StartCell As String, Cell As String)
    ConvCellToOY = GetCellOY(Cell) - GetCellOY(StartCell)
End Function



'===================================================

'Function GetCellX(Cell As String) As String
'    GetCellX = Mid(Cell, 1, 1)
'End Function
'
'Function GetCellY(Cell As String) As Integer
'    GetCellY = val(Mid(Cell, 2, 1))
'End Function
'
'Function ConvToCell(StartCellX As String, StartCellY As Integer, x As Integer, y As Integer)
'    ConvToCell = Chr(Asc(StartCellX) + x) + CStr(Round(StartCellY) + y)
'End Function
'
'Function ConvToOX(StartCellX As String, StartCellY As Integer, Cell As String)
'    ConvToOX = Asc(GetCellX(Cell)) - Asc(StartCellX)
'End Function
'
'Function ConvToOY(StartCellX As String, StartCellY As Integer, Cell As String)
'    ConvToOY = val(GetCellY(Cell)) - StartCellY
'End Function
'
'Function SetValByOXY(StartCell As String, x As Integer, y As Integer, value As String)
'    range(StartCell).Cells(y, x).value = value
'End Function
'
'
'Function ConvToCellRelCell(StartCell As String, x As Integer, y As Integer)
'    ConvToCellRelCell = ConvToCell(GetCellX(StartCell), GetCellY(StartCell), x, y)
'End Function
'
'Function SetColorCell(Cell As String, RGBColor As Long)
'    With range(Cell).Interior
'       .Color = RGBColor
'    End With
'End Function
'
'Function SetFontColorCell(Cell As String, RGBColor As Long)
'    With range(Cell).Font
'        .Color = RGBColor
'    End With
'End Function
