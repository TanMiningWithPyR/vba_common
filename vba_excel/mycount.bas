Attribute VB_Name = "mycount"
Public Function count(i As Integer, j As Integer, n As Integer, sheet As Worksheet, str As String) As Integer
    count = 0
    For x = i To i + n
    If sheet.Cells(x, j).Value = str Then
        count = count + 1
    End If
    Next
End Function
