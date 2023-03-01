Sub ScheduleGenerator()
    Dim sName As String
    sName = "??????"
    Dim i As Integer
    Dim j As Integer
    i = 0
    While (i < 75)
        j = 4
        Dim sCol As String
        Dim tmp As Integer
        tmp = i
        sCol = Num2Col(tmp)
        While (j < 24)
            Dim sCellName As String
            sCellName = sCol & CStr(j)
            Dim sCell As String
            sCell = Range(sCellName).Value
            If InStr(sCell, sName) <> 0 Then
                Range(sCellName).Value = "test"
            End If
            j = j + 1
        Wend
        i = i + 1

    Wend
End Sub
