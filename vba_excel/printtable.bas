Attribute VB_Name = "printtable"
Public Sub fillprint()

    Dim i2 As Integer
    i2 = 14
    For i = 15 To 214
        If Not IsEmpty(Sheet4.Cells(i, 3).Value) Then
            Sheet6.Cells(i2, 3).Value = Sheet4.Cells(i, 3).Value
            Sheet6.Cells(i2, 2).Value = Sheet4.Cells(i, 2).Value
            i2 = i2 + 1
        End If
    Next
    
    Call prset
    
End Sub
Private Sub prset()

   Dim iCount As Integer
   Dim a As Integer
   Dim b As Integer
   Dim MyPrintArea As String
   a = count(14, 3)
   b = 0
   
   iCount = a - b + 6
   MyPrintArea = "$A$1:$E$" & iCount
   Range("$A$1:$E$" & iCount).Select
   'Selection.Columns.AutoFit
   'Range("A8").Select
   'PrintTitlerow = "$6:$6"
   ActiveSheet.PageSetup.PrintTitleRows = ActiveSheet.Rows(13).Address
   ActiveSheet.PageSetup.PrintArea = MyPrintArea
   
End Sub
Private Function count(i As Integer, j As Integer) As Integer
    Dim n As Integer
    n = 0
    Do While Sheet6.Cells(i, j).Value <> ""
        i = i + 1
        n = n + 1
    Loop
    count = n
End Function
