Attribute VB_Name = "Mod_func"

Option Explicit

Public Function FindTextCell(ByVal ws As Worksheet, ByVal targetText As String) As Range

    Dim cell As Range
    Dim normalizedTarget As String

    normalizedTarget = NormalizeKeyText(targetText)

    For Each cell In ws.UsedRange.Cells
        If NormalizeKeyText(CStr(cell.value)) = normalizedTarget Then
            Set FindTextCell = cell
            Exit Function
        End If
    Next cell

End Function

Public Function NormalizeKeyText(ByVal textValue As String) As String
    NormalizeKeyText = UCase$(Trim$(textValue))
End Function

Public Function ReadCellToRight(ByVal anchorCell As Range) As String
    ReadCellToRight = Trim$(CStr(anchorCell.Offset(0, 1).value))
End Function

Public Function ReadRowValuesToRight(ByVal anchorCell As Range, Optional ByVal offsetRow As Long = 0, Optional ByVal offsetCol As Long = 1) As String()

    Dim ws As Worksheet
    Dim rowIndex As Long
    Dim startCol As Long
    Dim lastCol As Long
    Dim colIndex As Long
    Dim values() As String
    Dim cellValue As String

    Set ws = anchorCell.Worksheet
    rowIndex = anchorCell.row + offsetRow
    startCol = anchorCell.Column + offsetCol
    lastCol = ws.Cells(rowIndex, ws.Columns.Count).End(xlToLeft).Column

    For colIndex = startCol To lastCol
        cellValue = Trim$(CStr(ws.Cells(rowIndex, colIndex).value))
        If cellValue <> "" Then
            AppendToStringArray values, cellValue
        End If
    Next colIndex

    ReadRowValuesToRight = values

End Function

Public Function ReadColumnValuesBelow(ByVal anchorCell As Range, Optional ByVal offsetRow As Long = 0, Optional ByVal offsetCol As Long = 1, Optional ByVal lastRow As Long = 9) As String()

    Dim ws As Worksheet
    Dim dataCol As Long
    Dim startRow As Long
    Dim rowIndex As Long
    Dim values() As String
    Dim cellValue As String

    Set ws = anchorCell.Worksheet
    dataCol = anchorCell.Column + offsetCol
    startRow = anchorCell.row + offsetRow

    For rowIndex = startRow To lastRow
        cellValue = Trim$(CStr(ws.Cells(rowIndex, dataCol).value))
        If cellValue <> "" Then
            AppendToStringArray values, cellValue
        End If
    Next rowIndex

    ReadColumnValuesBelow = values

End Function

Public Sub AppendToStringArray(ByRef arr() As String, ByVal value As String)

    Dim nextIndex As Long

    If Not IsStringArrayAllocated(arr) Then
        ReDim arr(0 To 0)
        arr(0) = value
    Else
        nextIndex = UBound(arr) + 1
        ReDim Preserve arr(0 To nextIndex)
        arr(nextIndex) = value
    End If

End Sub

Public Function IsStringArrayAllocated(ByRef arr() As String) As Boolean

    Dim ub As Long

    On Error Resume Next
    ub = UBound(arr)
    IsStringArrayAllocated = (Err.Number = 0)
    On Error GoTo 0

End Function
