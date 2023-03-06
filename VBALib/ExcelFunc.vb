' convert the number to the cell columns name
Function Num2Col(n As Long)
    Dim res As Long
    res = n Mod 26
    Dim i As Long
    i = 65
    i = i + res
    Dim c As String
    c = Chr(i)
    Dim strPre As String
    strPre = ""
    If (n >= 26) Then
        n = (n - 26) \ 26
        strPre = Num2Col(n)
    End If
    strPre = strPre & c
    Num2Col = strPre
End Function


' locate cells within pointed words
Function CellsLocate(sUsedRange As Range, strName As String, nCellNum As Long, sCellPos() As Long)
    nCellNum = 0
    For Each Cell In sUsedRange.Cells
        If InStr(Cell.Value, strName) <> 0 Then
            sCellPos(nCellNum * 2) = Cell.Column
            sCellPos(nCellNum * 2 + 1) = Cell.Row
            nCellNum = nCellNum + 1
        End If
    Next
End Function


' Get Last Row Number in a Range
Function LastRowNumInRange(sRange As Range) As Long
    LastRowNumInRange = sRange.Row + sRange.Rows.Count - 1
End Function


' Get number of cells contained in sRange.
Function CellsCount(sRange As Range)
    CellsCount = sRange.Count
End Function


' Check if the cell is the first cell in a merged cell
Function IsFirstCellInMerge(sRange As Range)
    If sRange.MergeCells AND sRange.MergeArea.Cells(1).Address = sRange.Address Then
        IsFirstCellInMerge = True
    Else
        IsFirstCellInMerge = False
    End If
End Function


' Check if two cell in one Merged Cell
Function IsInOneCell(sCell1 As Range, sCell2 As Range)
    If NOT sCell1.MergeCells OR NOT sCell2.MergeCells Then
        IsInOneCell = False
    End If

    If sCell1.MergeArea.Cells(1).Address = sCell2.MergeArea.Cell(1).Address Then
        IsInOneCell = True
    Else
        IsInOneCell = False
    End If
End Function