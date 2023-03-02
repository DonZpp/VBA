Sub ScheduleGenerator()
    Dim strName As String
    strName = "ÔøÝ¼»Ô"
    Dim sCells(1000) As Long
    Dim nCellNum As Long
    nCellNum = 0
    Call CellsLocate(ActiveSheet.UsedRange, strName, nCellNum, sCells)

    Dim nRow As Long
    nRow = 4
    Dim nLastRow As Long
    nLast = LastRowNumInRange(ActiveSheet.UsedRange)

    For i = 0 To nCellNum - 1
        '  move to top
        ActiveSheet.Cells(nRow, sCells(i * 2)).Value = ActiveSheet.Cells(sCells(i * 2 + 1), sCells(i * 2)).Value
    Next

    '  delete cols without specific name
    Dim strColName As String
    For Each c In ActiveSheet.Rows(4).Cells
        If InStr(c.Value, strName) = 0 Then
            ' Call DeleteColumn(c.Column, xlToLeft)
        End If
    Next

    '  delete rows
    For j = nRow + 1 To nLast
        ' Call DeleteRow((j), (xlToUp))
    Next
End Sub