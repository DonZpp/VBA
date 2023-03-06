
Sub ScheduleGenerator()
    ActiveSheet.Range("3:3").AutoFilter

    Dim strName As String
    strName = "ÔøÝ¼»Ô"
    Dim sCells(1000) As Long
    Dim nCellNum As Long
    nCellNum = 0
    Call CellsLocate(ActiveSheet.UsedRange, strName, nCellNum, sCells)

    Dim nClassRow As Long
    nClassRow = 4

    '  move to top
    For i = 0 To nCellNum - 1
        Dim strNewName As String
        strNewName = ActiveSheet.Range("A" & CStr(sCells(i * 2 + 1))).Value
        strNewName = strNewName & ActiveSheet.Cells(sCells(i * 2 + 1), sCells(i * 2)).Value
        ActiveSheet.Cells(nClassRow, sCells(i * 2)).Value = strNewName
    Next

    '  delete cols without specific name
    Dim nRowCellCount As Long
    nRowCellCount = 0
    For Each sCellInRow In ActiveSheet.UsedRange.Range("3:3")
        If sCellInRow.Value = "" Then
            Exit For
        Else
            nRowCellCount = nRowCellCount + 1
        End If
    Next
    
    Dim nCurCol As Long
    nCurCol = 1
    Dim sCell As Range
    Dim strWeekDay As String
    strWeekDay = ""
    While nCurCol <= nRowCellCount
        Set sCell = ActiveSheet.Cells(nClassRow, nCurCol)
        If InStr(sCell.Value, strName) = 0 Then
            ' if there are weekday in 2nd row and it's right cell is empty, move It's value to right cell
            Dim sWeekDay As Range
            Set sWeekDay = ActiveSheet.Cells(2, nCurCol)
            Dim sWeekDayR As Range
            Set sWeekDayR = ActiveSheet.Cells(2, nCurCol + 1)
            
            If IsFirstCellInMerge(sWeekDay) And IsInOneCell(sWeekDay, sWeekDayR) Then
                strWeekDay = sWeekDay.Value
            End If

            ActiveSheet.Columns(nCurCol).Delete Shift:=xlToLeft
            nRowCellCount = nRowCellCount - 1
            If strWeekDay <> "" Then
                Dim tmpCell As Range
                Set tmpCell = ActiveSheet.Cells(2, nCurCol)
                tmpCell.Value = strWeekDay
                strWeekDay = ""
            End If
        Else
            nCurCol = nCurCol + 1
        End If
    Wend

    Dim nLastRow As Long
    nLastRow = LastRowNumInRange(ActiveSheet.UsedRange)
    ActiveSheet.Rows(CStr(5) & ":" & CStr(nLastRow)).Delete Shift:=xlToUp
    ActiveSheet.Rows("1:1").Delete Shift:=xlToUp

    ' change row height of class name 
    Range("3:3").RowHeight = 200
End Sub