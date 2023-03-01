' convert the number to the cell columns name
Function Num2Col(n As Integer)
    Dim res As Integer
    res = n Mod 26
    Dim i As Integer
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
Function CellsLocate(sUsedRange As Range, strName As String, sCellPos() As Long)
    Dim pos As Integer
    pos = 0
    For Each Cell In sUsedRange.Cells
        If InStr(Cell.Value, strName) <> 0 Then
            sCellPos(pos * 2) = Cell.Column
            sCellPos(pos * 2 + 1) = Cell.Row
            pos = pos + 1
        End If
    Next
End Function