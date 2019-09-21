Attribute VB_Name = "findij"
'��һ���ض���sheet���ض���������(��cell ��i ��j����ʼ�����������ֹͣ��������һ����ֵ(long type)�������ظ���ֵ���ڵ��к�
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
'��һ���ض���sheet���ض���������(��cell ��i ��j����ʼ�����������ֹͣ��������һ����ֵ(long type)�������ظ���ֵ���ڵ��к�
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
 '��һ���ض���sheet���ض���������(��cell ��i ��j����ʼ�����������ֹͣ��������һ���ַ����������ظ��ַ������ڵ��к�
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
 '��һ���ض���sheet���ض���������(��cell ��i ��j����ʼ�����������ֹͣ��������һ���ַ����������ظ��ַ������ڵ��к�
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
 '��һ���ض���sheet���ض���������(��cell ��i ��j����ʼ����n����������һ���ַ����������ظ��ַ������ڵ��к�
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
 '��һ���ض���sheet���ض���������(��cell ��i ��j����ʼ����n����������һ���ַ����������ظ��ַ������ڵ��к�
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
