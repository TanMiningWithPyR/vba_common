Attribute VB_Name = "mycopy"
'从sheet1某一个单元格（i1,j1）开始向下复制列到（i2,j2),共有n个
Public Sub copycolumn_n(Sheet1 As Worksheet, sheet2 As Table, i1 As Integer, j1 As Integer, i2 As Integer, j2 As Integer, n As Integer)

     Do While n <> 0
        sheet2.Cell(i2, j2).Shape.TextFrame.TextRange.Text = Sheet1.Cells(i1, j1).Value
        i1 = i1 + 1
        i2 = i2 + 1
        n = n - 1
     Loop
    
End Sub

'从sheet1某一个单元格（i1,j1）开始向右下复制一块区域到（i2,j2),n1为一共几行，n2为一共几列
Public Sub copyerea_n(Sheet1 As Worksheet, sheet2 As Table, i1 As Integer, j1 As Integer, i2 As Integer, j2 As Integer, n1 As Integer, n2 As Integer)
        
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
