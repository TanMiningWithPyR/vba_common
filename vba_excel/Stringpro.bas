Attribute VB_Name = "Stringpro"
Public Function GetstrToasc(name As String, i As Integer) As Integer 'ȡ���ַ�����i���ַ�����תΪ��Ӧ��ASCII��

    Dim getstr As String
    getstr = Mid(name, i, 1)
    GetstrToasc = Asc(getstr)
    
End Function
Public Function Get_num(name1 As String) As Integer '�ж�һ���ַ��������Ƿ���".",���ҷ������һ�γ������ַ��������ǵڼ����ַ�

    Dim Is_y As Integer
    Dim j As Integer
    j = 1
    Do
        Is_y = GetstrToasc(name1, j)
        j = j + 1
    Loop While Is_y <> 46
    
    Get_num = j - 1
        
End Function
Public Function splitatnote(SplitString As String, note As String) As String
    splitatnote = Left(SplitString, InStr(SplitString, note) - 1)
End Function


