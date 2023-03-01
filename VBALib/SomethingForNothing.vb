Sub ScheduleGenerator()
    Dim strName As String
    strName = "ÔøÝ¼»Ô"
    Dim sCells(100) As Variant
    Call CellsLocate(ActiveSheet.UsedRange, strName, sCells)
    For Each sCellPos In sCells
        MsgBox sCellPos
    Next
End Sub