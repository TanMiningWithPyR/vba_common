Attribute VB_Name = "findij"
'在一个特定的sheet，特定的列里面(从cell （i ，j）开始，如果遇到空停止），查找一个数值(long type)，并返回该数值所在的行号
Public Function IntFindrow(Number As Long, Sheet As Worksheet, i As Integer, j As Integer) As Integer
           
    Do While Not IsEmpty(Sheet.Cells(i, j).Value)
        If Sheet.Cells(i, j).Value = Number Then
            IntFindrow = i
            Exit Do
        Else
            i = i + 1
            If IsEmpty(Sheet.Cells(i, j).Value) Then
                MsgBox ("no this number")
            End If
        End If
    Loop

End Function
'在一个特定的sheet，特定的行里面(从cell （i ，j）开始，如果遇到空停止），查找一个数值(long type)，并返回该数值所在的列号
Public Function IntFindcolumn(Number As Long, Sheet As Worksheet, i As Integer, j As Integer) As Integer
      
    Do While Not IsEmpty(Sheet.Cells(i, j).Value)
        If Sheet.Cells(i, j).Value = Number Then
            IntFindcolumn = j
            Exit Do
        Else
            j = j + 1
            If IsEmpty(Sheet.Cells(i, j).Value) Then
                MsgBox ("no this number")
            End If
        End If
    Loop

End Function
 '在一个特定的sheet，特定的列里面(从cell （i ，j）开始，如果遇到空停止），查找一个字符串，并返回该字符串所在的行号
Public Function IntFind_str_row(name As String, Sheet As Worksheet, i As Integer, j As Integer) As Integer
           
    Do While Not IsEmpty(Sheet.Cells(i, j).Value)
        If Sheet.Cells(i, j).Value = name Then
            IntFind_str_row = i
            Exit Do
        Else
            i = i + 1
            If IsEmpty(Sheet.Cells(i, j).Value) Then
                MsgBox ("no this string")
            End If
        End If
    Loop

End Function
 '在一个特定的sheet，特定的行里面(从cell （i ，j）开始，如果遇到空停止），查找一个字符串，并返回该字符串所在的列号
Public Function IntFind_str_col(name As String, Sheet As Worksheet, i As Integer, j As Integer) As Integer
           
    Do While Not IsEmpty(Sheet.Cells(i, j).Value)
        If Sheet.Cells(i, j).Value = name Then
            IntFind_str_col = j
            Exit Do
        Else
            j = j + 1
            If IsEmpty(Sheet.Cells(i, j).Value) Then
                MsgBox ("no this string")
            End If
        End If
    Loop

End Function
 '在一个特定的sheet，特定的列里面(从cell （i ，j）开始，找n个），查找一个字符串，并返回该字符串所在的行号
Public Function IntFind_str_row_n(name As String, Sheet As Worksheet, i As Integer, j As Integer, n As Integer) As Integer
           
    For x = 1 To n
        If Sheet.Cells(i, j).Value = name Then
            IntFind_str_row_n = i
            Exit For
        Else
            i = i + 1
            If x = n Then
                MsgBox ("no this string")
            End If
        End If
    Next

End Function
 '在一个特定的sheet，特定的行里面(从cell （i ，j）开始，找n个），查找一个字符串，并返回该字符串所在的列号
Public Function IntFind_str_col_n(name As String, Sheet As Worksheet, i As Integer, j As Integer, n As Integer) As Integer
           
    For x = 1 To n
        If Sheet.Cells(i, j).Value = name Then
            IntFind_str_col_n = j
            Exit For
        Else
            j = j + 1
            If x = n Then
                MsgBox ("no this string")
            End If
        End If
    Next

End Function
