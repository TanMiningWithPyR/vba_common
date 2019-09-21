Attribute VB_Name = "mycopy"
'��sheet1ĳһ����Ԫ��i1,j1����ʼ���¸����е���i2,j2),�����ո�ֹͣ
Public Sub copycolumn(Sheet1 As Worksheet, sheet2 As Worksheet, i1 As Integer, j1 As Integer, i2 As Integer, j2 As Integer)

     Do While Not IsEmpty(Sheet1.Cells(i1, j1).Value)
        sheet2.Cells(i2, j2).Value = Sheet1.Cells(i1, j1).Value
        i1 = i1 + 1
        i2 = i2 + 1
     Loop
    
End Sub

'��sheet1ĳһ����Ԫ��i1,j1����ʼ���Ҹ����е���i2,j2)�������ո�ֹͣ
Public Sub copyrow(Sheet1 As Worksheet, sheet2 As Worksheet, i1 As Integer, j1 As Integer, i2 As Integer, j2 As Integer)

     Do While Not IsEmpty(Sheet1.Cells(i1, j1).Value)
        sheet2.Cells(i2, j2).Value = Sheet1.Cells(i1, j1).Value
        j1 = j1 + 1
        j2 = j2 + 1
     Loop
    
End Sub
'��sheet1ĳһ����Ԫ��i1,j1����ʼ���Ҹ���һ�����򵽣�i2,j2),nΪһ���м��У����������ո�ֹͣ
Public Sub copyerea(Sheet1 As Worksheet, sheet2 As Worksheet, i1 As Integer, j1 As Integer, i2 As Integer, j2 As Integer, n As Integer)
        Dim a, b As Integer
        a = i1
        b = i2
        
        For c = 1 To n
                i1 = a
                i2 = b
            Call copycolumn(Sheet1, sheet2, i1, j1, i2, j2)
                j1 = j1 + 1
                j2 = j2 + 1
        Next
        
End Sub
'��sheet1ĳһ����Ԫ��i1,j1����ʼ���¸����е���i2,j2),����n��
Public Sub copycolumn_n(Sheet1 As Worksheet, sheet2 As Worksheet, i1 As Integer, j1 As Integer, i2 As Integer, j2 As Integer, n As Integer)

     Do While n <> 0
        sheet2.Cells(i2, j2).Value = Sheet1.Cells(i1, j1).Value
        i1 = i1 + 1
        i2 = i2 + 1
        n = n - 1
     Loop
    
End Sub
'��sheet1ĳһ����Ԫ��i1,j1����ʼ���Ҹ����е���i2,j2)������n��
Public Sub copyrow_n(Sheet1 As Worksheet, sheet2 As Worksheet, i1 As Integer, j1 As Integer, i2 As Integer, j2 As Integer, n As Integer)

     Do While n <> 0
        sheet2.Cells(i2, j2).Value = Sheet1.Cells(i1, j1).Value
        j1 = j1 + 1
        j2 = j2 + 1
        n = n - 1
     Loop
    
End Sub
'��sheet1ĳһ����Ԫ��i1,j1����ʼ�����¸���һ�����򵽣�i2,j2),n1Ϊһ�����У�n2Ϊһ������
Public Sub copyerea_n(Sheet1 As Worksheet, sheet2 As Worksheet, i1 As Integer, j1 As Integer, i2 As Integer, j2 As Integer, n1 As Integer, n2 As Integer)
        
        Dim row1 As Integer
        Dim row2 As Integer
        Dim n As Integer
        
    For c = 1 To n2
        If c = 1 Then
        row1 = i1
        row2 = i2
        n = n1
        End If
        i1 = row1
        i2 = row2
        n1 = n
        Call copycolumn_n(Sheet1, sheet2, i1, j1 + c - 1, i2, j2 + c - 1, n1)
    Next
        
End Sub
Public Sub remove_duplicates(sheet As Worksheet, j As Integer, Optional i As Integer = 1, Optional n As Integer = 300) 'ȥ�����}�Ć�Ԫ��j ���У�i��ĵڎ����_ʼ�����һ���Ўׂ���Ԫ��
     Dim t1 As Integer
     Dim t2 As Integer
     For t1 = i To i + 299 '����һ�����ñ�
        If Not IsEmpty(sheet.Cells(t1, j).Value) Then
            For t2 = t1 + 1 To i + 300
                If sheet.Cells(t2, j).Value = sheet.Cells(t1, j).Value Then
                    sheet.Cells(t2, j).Value = ""
                End If
            Next
        End If
    Next
End Sub
'�Ԅ������һ�еă���
Public Sub autofill(sheet As Worksheet, j As Integer, i2 As Integer, Optional i1 As Integer = 1)
Dim n As Integer
    For n = i1 + 1 To i2
        If IsEmpty(sheet.Cells(n, j).Value) Then
            sheet.Cells(n, j).Value = sheet.Cells(n - 1, j).Value
        End If
    Next
End Sub
'����͸�ӱ���飬Ȼ���Ƴɱ���ж��ĵط��Ǳ�ͷ
Public Sub transport_by(sheet As Worksheet, i1 As Integer, i2 As Integer, j As Integer)
    Dim row As Integer
    Dim column As Integer
     row = 1
     column = 0
    For n = i1 To i2
        If IsNumeric(sheet.Cells(n, j).Value) Then
            row = 1
            column = column + 1
            sheet.Cells(row, column).Value = sheet.Cells(n, j).Value
            row = row + 1
        Else
            sheet.Cells(row, column).Value = sheet.Cells(n, j).Value
            row = row + 1
        End If
    Next
End Sub
'�����ո񣬾�ɾ����
Public Sub row_blank_delete(sheet As Worksheet, i1 As Integer, i2 As Integer, j As Integer)
n = 1
For i = i1 To i2
    If IsEmpty(sheet.Cells(i, j).Value) Then
        sheet.Rows(i).Delete
        n = n + 1
        i = i - 1
    Else
        n = n + 1
    End If
    
    If n = i2 - i1 Then
        MsgBox ("finished")
        Exit Sub
    End If
Next
End Sub
