' 将数字转换为26进制，且符号为A-Z
Function Num2Col(n As Integer)
    Dim res As Integer
    res = n Mod 26
    Dim i As Integer
    i = 65
    i = i + res
    Dim c As String
    c = Chr(i)
    Dim sPre As String
    sPre = ""
    If (n >= 26) Then
        n = (n - 26) \ 26
        sPre = Num2Col(n)
    End If
    sPre = sPre & c
    Num2Col = sPre
End Function
