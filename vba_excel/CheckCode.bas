Attribute VB_Name = "CheckCode"
'用 i1行，j列，共i2-i1行数据生成校验码,最多可有60个0，1量
Public Function GenerateCCode(sheet As Worksheet, i1 As Integer, i2 As Integer, j As Integer) As String
        Dim a(5) As Integer
        Dim b(9) As Integer
        Dim n As Integer
        Dim y As Integer
        
            GenerateCCode = "-"
            n = 0
            For x = i1 To i2
                For y = 0 To 5
                    If Not IsEmpty(sheet.Cells(x, j).Value) Then
                        a(y) = sheet.Cells(x, j).Value
                    Else
                        y = y - 1
                    End If
                    
                    If x = i2 Then
                    x = x + 1
                    Exit For
                    End If
                    
                    x = x + 1
                Next
                b(n) = thirtytwo(a(0), a(1), a(2), a(3), a(4), a(5))
                GenerateCCode = GenerateCCode + Hex2(b(n))
                n = n + 1
                x = x - 1
                a(0) = 0
                a(1) = 0
                a(2) = 0
                a(3) = 0
                a(4) = 0
                a(5) = 0
            Next
End Function
Public Function Hex2(n As Integer) As String
    If n < 10 And n >= 0 Then
        Hex2 = Chr(n + 48)
    ElseIf n >= 10 And n < 36 Then
        Hex2 = Chr(n + 55)
    ElseIf n >= 36 And n < 62 Then
        Hex2 = Chr(n + 61)
    ElseIf n = 62 Or n = 63 Then
        Hex2 = Chr(n + 1)
    Else
        MsgBox ("wrong number")
    End If
End Function
Public Function thirtytwo(a As Integer, Optional b As Integer = 0, Optional c As Integer = 0, Optional d As Integer = 0, Optional e As Integer = 0, Optional f As Integer) As Integer
    thirtytwo = a * 1 + b * 2 + c * 4 + d * 8 + e * 16 + f * 32
End Function

