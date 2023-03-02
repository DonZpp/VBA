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

' Delete a column
Function DeleteColumn(nCol As Long, xlTo As Long)
    Dim strColName As String
    strColName = Num2Col(nCol)
    Columns(strColName & ":" & strColName).Select
    Selection.Delete Shift:=xlTo
End Function

' Delete a row
Function DeleteRow(nRow As Long , xlTo As Long)
    Dim strRow As String
    strRow = CStr(nRow)
    Rows(strRow & ":" & strRow).Select
    Selection.Delete Shift := xlTo
End Function