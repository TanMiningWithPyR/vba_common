Attribute VB_Name = "Stringpro"
Public Function GetstrToasc(name As String, i As Integer) As Integer '取出字符串第i个字符，并转为对应的ASCII码

    Dim getstr As String
    getstr = Mid(name, i, 1)
    GetstrToasc = Asc(getstr)
    
End Function
Public Function Get_num(name1 As String) As Integer '判断一个字符串里面是否有".",并且返回其第一次出现在字符串里面是第几个字符

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


