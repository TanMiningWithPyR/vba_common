Attribute VB_Name = "insertpicture"
Public Sub delete_picture_page(wkb As Workbook)
With wkb
    Dim dnr(10) As Integer
        dnr(1) = delete_picture_row(338, .Sheets("parameter").Cells(2, 10).Value, wkb)
        dnr(10) = dnr(1)
        dnr(2) = delete_picture_row(466 - dnr(10), .Sheets("parameter").Cells(3, 10).Value, wkb)
        dnr(10) = dnr(10) + dnr(2)
        dnr(3) = delete_picture_row(583 - dnr(10), .Sheets("parameter").Cells(4, 10).Value, wkb)
        dnr(10) = dnr(10) + dnr(3)
        dnr(4) = delete_picture_row(689 - dnr(10), .Sheets("parameter").Cells(5, 10).Value, wkb)
        dnr(10) = dnr(10) + dnr(4)
        dnr(5) = delete_picture_row(784 - dnr(10), .Sheets("parameter").Cells(6, 10).Value, wkb)
        dnr(10) = dnr(10) + dnr(5)
        dnr(6) = delete_picture_row(890 - dnr(10), .Sheets("parameter").Cells(7, 10).Value, wkb)
        dnr(10) = dnr(10) + dnr(6)
        dnr(7) = delete_picture_row(996 - dnr(10), .Sheets("parameter").Cells(8, 10).Value, wkb)
        dnr(10) = dnr(10) + dnr(7)
        dnr(8) = delete_picture_row(1091 - dnr(10), .Sheets("parameter").Cells(9, 10).Value, wkb)
        dnr(10) = dnr(10) + dnr(8)
End With
End Sub
Public Function delete_picture_row(i As Integer, n As Integer, wkb As Workbook) As Integer 'i 为模块照片的末行, n 为模块照片数量 ,返回删除的行数
With wkb
    Dim p, q, np, id As Integer
        p = n \ 4
        q = n Mod 4
    If q = 0 Then
        np = p
    Else
        np = p + 1
    End If
        dnr = (6 - np) * 14  'dnr 为需要被删除的行数
        If dnr <> 0 Then
            .Sheets("report").Rows(i - dnr + 1 & ":" & i).Delete
        End If
        delete_picture_row = dnr
End With
End Function
Public Sub insertpicture(dealer As String, wkb As Workbook)
With wkb
.Sheets("report").Activate
Dim i As Integer
i = 2
j1 = 4
j2 = 4
j3 = 4
j4 = 4
j5 = 4
j6 = 4
j7 = 4
j8 = 4
'统计每个模块的照片数量
Dim picturecount(10) As Integer
For ip = 0 To 10
picturecount(ip) = 0
Next

ireport1 = 260
ireport2 = 388
ireport3 = 505
ireport4 = 611
ireport5 = 706
ireport6 = 812
ireport7 = 918
ireport8 = 1013
Do While Not IsEmpty(.Sheets("parameter").Cells(i, 1).Value)
    If .Sheets("parameter").Cells(i, 2).Value = 1 Then
        picturecount(1) = picturecount(1) + 1
        If j1 > 11 Then
            If picturecount(1) Mod 4 = 1 Then
                ireport1 = ireport1 + 9
                j1 = 4
            Else
                ireport1 = ireport1 + 5
                j1 = 4
            End If
        End If
        .Sheets("report").Cells(ireport1 - 1, j1).Value = .Sheets("parameter").Cells(i, 3).Value
        Call InsertPictureInRange(Application.ThisWorkbook.Path & "\picture\pfile_" & dealer & "\" & .Sheets("parameter").Cells(i, 1).Value, .Sheets("report").Range(Cells(ireport1, j1), Cells(ireport1, j1)))
        j1 = j1 + 7
    End If
    
    If .Sheets("parameter").Cells(i, 2).Value = 2 Then
        picturecount(2) = picturecount(2) + 1
        If j2 > 11 Then
            If picturecount(2) Mod 4 = 1 Then
                ireport2 = ireport2 + 9
                j2 = 4
            Else
                ireport2 = ireport2 + 5
                j2 = 4
            End If
        End If
        .Sheets("report").Cells(ireport2 - 1, j2).Value = .Sheets("parameter").Cells(i, 3).Value
        Call InsertPictureInRange(Application.ThisWorkbook.Path & "\picture\pfile_" & dealer & "\" & .Sheets("parameter").Cells(i, 1).Value, .Sheets("report").Range(Cells(ireport2, j2), Cells(ireport2, j2)))
        j2 = j2 + 7
    End If
    
    If .Sheets("parameter").Cells(i, 2).Value = 3 Then
        picturecount(3) = picturecount(3) + 1
        If j3 > 11 Then
            If picturecount(3) Mod 4 = 1 Then
                ireport3 = ireport3 + 9
                j3 = 4
            Else
                ireport3 = ireport3 + 5
                j3 = 4
            End If
        End If
        .Sheets("report").Cells(ireport3 - 1, j3).Value = .Sheets("parameter").Cells(i, 3).Value
        Call InsertPictureInRange(Application.ThisWorkbook.Path & "\picture\pfile_" & dealer & "\" & .Sheets("parameter").Cells(i, 1).Value, .Sheets("report").Range(Cells(ireport3, j3), Cells(ireport3, j3)))
        j3 = j3 + 7
    End If
    
    If .Sheets("parameter").Cells(i, 2).Value = 4 Then
        picturecount(4) = picturecount(4) + 1
        If j4 > 11 Then
            If picturecount(4) Mod 4 = 1 Then
                ireport4 = ireport4 + 9
                j4 = 4
            Else
                ireport4 = ireport4 + 5
                j4 = 4
            End If
        End If
        .Sheets("report").Cells(ireport4 - 1, j4).Value = .Sheets("parameter").Cells(i, 3).Value
        Call InsertPictureInRange(Application.ThisWorkbook.Path & "\picture\pfile_" & dealer & "\" & .Sheets("parameter").Cells(i, 1).Value, .Sheets("report").Range(Cells(ireport4, j4), Cells(ireport4, j4)))
        j4 = j4 + 7
    End If
    
    If .Sheets("parameter").Cells(i, 2).Value = 5 Then
        picturecount(5) = picturecount(5) + 1
        If j5 > 11 Then
            If picturecount(5) Mod 4 = 1 Then
                ireport5 = ireport5 + 9
                j5 = 4
            Else
                ireport5 = ireport5 + 5
                j5 = 4
            End If
        End If
        .Sheets("report").Cells(ireport5 - 1, j5).Value = .Sheets("parameter").Cells(i, 3).Value
        Call InsertPictureInRange(Application.ThisWorkbook.Path & "\picture\pfile_" & dealer & "\" & .Sheets("parameter").Cells(i, 1).Value, .Sheets("report").Range(Cells(ireport5, j5), Cells(ireport5, j5)))
        j5 = j5 + 7
    End If
    
    If .Sheets("parameter").Cells(i, 2).Value = 6 Then
        picturecount(6) = picturecount(6) + 1
        If j6 > 11 Then
            If picturecount(6) Mod 4 = 1 Then
                ireport6 = ireport6 + 9
                j6 = 4
            Else
                ireport6 = ireport6 + 5
                j6 = 4
            End If
        End If
         .Sheets("report").Cells(ireport6 - 1, j6).Value = .Sheets("parameter").Cells(i, 3).Value
        Call InsertPictureInRange(Application.ThisWorkbook.Path & "\picture\pfile_" & dealer & "\" & .Sheets("parameter").Cells(i, 1).Value, .Sheets("report").Range(Cells(ireport6, j6), Cells(ireport6, j6)))
        j6 = j6 + 7
    End If
   
    If .Sheets("parameter").Cells(i, 2).Value = 7 Then
        picturecount(7) = picturecount(7) + 1
        If j7 > 11 Then
            If picturecount(7) Mod 4 = 1 Then
                ireport7 = ireport7 + 9
                j7 = 4
            Else
                ireport7 = ireport7 + 5
                j7 = 4
            End If
        End If
         .Sheets("report").Cells(ireport7 - 1, j7).Value = .Sheets("parameter").Cells(i, 3).Value
        Call InsertPictureInRange(Application.ThisWorkbook.Path & "\picture\pfile_" & dealer & "\" & .Sheets("parameter").Cells(i, 1).Value, .Sheets("report").Range(Cells(ireport7, j7), Cells(ireport7, j7)))
        j7 = j7 + 7
    End If
    
    If .Sheets("parameter").Cells(i, 2).Value = 8 Then
        picturecount(8) = picturecount(8) + 1
        If j8 > 11 Then
            If picturecount(8) Mod 4 = 1 Then
                ireport8 = ireport8 + 9
                j8 = 4
            Else
                ireport8 = ireport8 + 5
                j8 = 4
            End If
        End If
         .Sheets("report").Cells(ireport8 - 1, j8).Value = .Sheets("parameter").Cells(i, 3).Value
        Call InsertPictureInRange(Application.ThisWorkbook.Path & "\picture\pfile_" & dealer & "\" & .Sheets("parameter").Cells(i, 1).Value, .Sheets("report").Range(Cells(ireport8, j8), Cells(ireport8, j8)))
        j8 = j8 + 7
    End If
    i = i + 1
Loop
For i = 1 To 8
    .Sheets("parameter").Cells(i + 1, 10).Value = picturecount(i)
Next
End With
End Sub
Sub InsertPictureInRange(PictureFileName As String, TargetCells As Range)
' inserts a picture and resizes it to fit the TargetCells range
Dim p As Object, t As Double, l As Double, w As Double, h As Double
    If TypeName(ActiveSheet) <> "Worksheet" Then Exit Sub
    If Dir(PictureFileName) = "" Then Exit Sub
    ' import picture
    Set p = ActiveSheet.Pictures.insert(PictureFileName)
    ' determine positions
    With TargetCells
        t = .Top
        l = .Left
        w = .Offset(0, .Columns.Count).Left - .Left
        h = .Offset(.Rows.Count, 0).Top - .Top
    End With
    ' position picture
    p.Placement = xlMoveAndSize '设置图片可以随单元格的变动而改变大小和位置
    p.ShapeRange.LockAspectRatio = msoFalse '取消图片纵横比锁定
    With p
        .Top = t
        .Left = l
        .Width = w
        .Height = h
    End With
    Set p = Nothing
End Sub
