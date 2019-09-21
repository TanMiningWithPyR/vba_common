Attribute VB_Name = "vector"
'����nά��������������,����Double��
Public Function DoubleRow_multi(sheet1 As Worksheet, sheet2 As Worksheet, i1 As Integer, j1 As Integer, i2 As Integer, j2 As Integer, n As Integer) As Double
    Dim b As Double
    b = 0
    DoubleRow_multi = 0
    
    For a = 1 To n
        b = sheet1.Cells(i1, j1).Value * sheet2.Cells(i2, j2).Value
        DoubleRow_multi = DoubleRow_multi + b
        j1 = j1 + 1
        j2 = j2 + 1
    Next
    
End Function
    
'����nά��������������,����Double��
Public Function DoubleColumn_multi(sheet1 As Worksheet, sheet2 As Worksheet, i1 As Integer, j1 As Integer, i2 As Integer, j2 As Integer, n As Integer) As Double
    Dim b As Double
    b = 0
    DoubleColumn_multi = 0
    
    For a = 1 To n
        b = sheet1.Cells(i1, j1).Value * sheet2.Cells(i2, j2).Value
        DoubleColumn_multi = DoubleColumn_multi + b
        i1 = i1 + 1
        i2 = i2 + 1
    Next
    
End Function

'����ת��
Public Sub Vectortranspose(sheet1 As Worksheet, sheet2 As Worksheet, i1 As Integer, j1 As Integer, i2 As Integer, j2 As Integer, n As Integer, direction As Boolean)
                
        If direction = True Then '��������������
            For a = 1 To n
                sheet2.Cells(i2, j2).Value = sheet1.Cells(i1, j1).Value
                j1 = j1 + 1
                i2 = i2 + 1
            Next
        Else
            For a = 1 To n
                sheet2.Cells(i2, j2).Value = sheet1.Cells(i1, j1).Value
                i1 = i1 + 1
                j2 = j2 + 1
                
            Next
        End If
      
        
End Sub


