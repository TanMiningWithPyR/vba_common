Attribute VB_Name = "doby"
'合并同类项
Public Sub mergeby(sheetby As Worksheet, i_by As Integer, j_by As Integer, n_by As Integer, sheetmerge As Worksheet, i_merge As Integer, j_merge As Integer, n_merge As Integer)
    Dim i As Integer
    i = 0
    n_merge = 0
    For i = i_by To i_by + n_by - 1
        sheetmerge.Cells(i_merge + n_merge, j_merge).Value = sheetby.Cells(i, j_by).Value
        i = i + 1
        Do While sheetmerge.Cells(i_merge + n_merge, j_merge).Value = sheetby.Cells(i, j_by).Value And i < i_by + n_by
                i = i + 1
        Loop
        i = i - 1
        n_merge = n_merge + 1
    Next
End Sub
'合并同类项，同类项以空格形式出现
Public Sub mergeblankby(sheetby As Worksheet, i_by As Integer, j_by As Integer, n_by As Integer, sheetmerge As Worksheet, i_merge As Integer, j_merge As Integer, n_merge As Integer)
   Dim i As Integer
   i = 0
   n_merge = 0
     For i = i_by To i_by + n_by - 1
        sheetmerge.Cells(i_merge + n_merge, j_merge).Value = sheetby.Cells(i, j_by).Value
        i = i + 1
        If i > i_by + n_by Then
            Exit For
        End If
        Do While IsEmpty(sheetby.Cells(i, j_by).Value) And i < i_by + n_by
                i = i + 1
                        If i > i_by + n_by Then
                            Exit For
                        End If
        Loop
        i = i - 1
        n_merge = n_merge + 1
    Next
End Sub
Private Sub sumby(sheetby As Worksheet, i_by As Integer, j_by As Integer, n_by As Integer, i_do As Integer, j_do As Integer, sheetsum As Worksheet, i_merge As Integer, j_merge As Integer)
    Dim i_sum, j_sum As Integer
    i_sum = i_merge
    j_sum = j_merge + j_do - j_by
    Dim i, n_merge As Integer
    i = 0
    n_merge = 0
    For i = i_by To i_by + n_by - 1
        sheetsum.Cells(i_sum + n_merge, j_sum).Value = sheetby.Cells(i, j_do).Value
        i = i + 1
        Do While sheetsum.Cells(i_merge + n_merge, j_merge).Value = sheetby.Cells(i, j_by).Value And i < i_by + n_by
                sheetsum.Cells(i_sum + n_merge, j_sum).Value = sheetby.Cells(i, j_do).Value + sheetsum.Cells(i_sum + n_merge, j_sum).Value
                i = i + 1
        Loop
        i = i - 1
        n_merge = n_merge + 1
    Next
End Sub
Private Sub productby(sheetby As Worksheet, i_by As Integer, j_by As Integer, n_by As Integer, i_do As Integer, j_do As Integer, sheetproduct As Worksheet, i_merge As Integer, j_merge As Integer)
    Dim i_product, j_product As Integer
    i_product = i_merge
    j_product = j_merge + j_do - j_by
    Dim i, n_merge As Integer
    i = 0
    n_merge = 0
    For i = i_by To i_by + n_by - 1
        sheetproduct.Cells(i_product + n_merge, j_product).Value = sheetby.Cells(i, j_do).Value
        i = i + 1
        Do While sheetproduct.Cells(i_merge + n_merge, j_merge).Value = sheetby.Cells(i, j_by).Value And i < i_by + n_by
                sheetproduct.Cells(i_product + n_merge, j_product).Value = sheetby.Cells(i, j_do).Value * sheetproduct.Cells(i_product + n_merge, j_product).Value
                i = i + 1
        Loop
        i = i - 1
        n_merge = n_merge + 1
    Next
End Sub
Private Sub sumblankby(sheetby As Worksheet, i_by As Integer, j_by As Integer, n_by As Integer, i_do As Integer, j_do As Integer, sheetsum As Worksheet, i_merge As Integer, j_merge As Integer)
   Dim i_sum, j_sum As Integer
   i_sum = i_merge
   j_sum = j_merge + j_do - j_by
   Dim i, n_merge As Integer
   i = 0
   n_merge = 0
     For i = i_by To i_by + n_by - 1
        sheetsum.Cells(i_sum + n_merge, j_sum).Value = sheetby.Cells(i, j_do).Value
        i = i + 1
        If i > i_by + n_by Then
            Exit For
        End If
        Do While IsEmpty(sheetby.Cells(i, j_by).Value) And i < i_by + n_by
                sheetsum.Cells(i_sum + n_merge, j_sum).Value = sheetby.Cells(i, j_do).Value + sheetsum.Cells(i_sum + n_merge, j_sum).Value
                i = i + 1
                        If i > i_by + n_by Then
                            Exit For
                        End If
        Loop
        i = i - 1
        n_merge = n_merge + 1
    Next
End Sub
Private Sub productblankby(sheetby As Worksheet, i_by As Integer, j_by As Integer, n_by As Integer, i_do As Integer, j_do As Integer, sheetproduct As Worksheet, i_merge As Integer, j_merge As Integer)
   Dim i_product, j_product As Integer
   i_product = i_merge
   j_product = j_merge + j_do - j_by
   Dim i, n_merge As Integer
   i = 0
   n_merge = 0
     For i = i_by To i_by + n_by - 1
        sheetproduct.Cells(i_product + n_merge, j_product).Value = sheetby.Cells(i, j_do).Value
        i = i + 1
        If i > i_by + n_by Then
            Exit For
        End If
        Do While IsEmpty(sheetby.Cells(i, j_by).Value) And i < i_by + n_by
                sheetproduct.Cells(i_product + n_merge, j_product).Value = sheetby.Cells(i, j_do).Value * sheetproduct.Cells(i_product + n_merge, j_product).Value
                i = i + 1
                        If i > i_by + n_by Then
                            Exit For
                        End If
        Loop
        i = i - 1
        n_merge = n_merge + 1
    Next
End Sub
Public Sub sumareaby(sheetby As Worksheet, i_by As Integer, j_by As Integer, n_by As Integer, i_do As Integer, j_do As Integer, j_do_n As Integer, sheetsum As Worksheet, i_merge As Integer, j_merge As Integer)
    Dim j As Integer
    Call mergeby(sheetby, i_by, j_by, n_by, sheetsum, i_merge, j_merge, 0)
    For j = j_do To j_do + j_do_n - 1
        Call sumby(sheetby, i_by, j_by, n_by, i_do, j, sheetsum, i_merge, j_merge)
    Next
End Sub
Public Sub productareaby(sheetby As Worksheet, i_by As Integer, j_by As Integer, n_by As Integer, i_do As Integer, j_do As Integer, j_do_n As Integer, sheetproduct As Worksheet, i_merge As Integer, j_merge As Integer)
    Dim j As Integer
    Call mergeby(sheetby, i_by, j_by, n_by, sheetproduct, i_merge, j_merge, 0)
    For j = j_do To j_do + j_do_n - 1
        Call productby(sheetby, i_by, j_by, n_by, i_do, j, sheetproduct, i_merge, j_merge)
    Next
End Sub
Public Sub sumareablankby(sheetby As Worksheet, i_by As Integer, j_by As Integer, n_by As Integer, i_do As Integer, j_do As Integer, j_do_n As Integer, sheetsum As Worksheet, i_merge As Integer, j_merge As Integer)
    Dim j As Integer
    Call mergeblankby(sheetby, i_by, j_by, n_by, sheetsum, i_merge, j_merge, 0)
    For j = j_do To j_do + j_do_n - 1
        Call sumblankby(sheetby, i_by, j_by, n_by, i_do, j, sheetsum, i_merge, j_merge)
    Next
End Sub
Public Sub productareablankby(sheetby As Worksheet, i_by As Integer, j_by As Integer, n_by As Integer, i_do As Integer, j_do As Integer, j_do_n As Integer, sheetproduct As Worksheet, i_merge As Integer, j_merge As Integer)
    Dim j As Integer
    Call mergeblankby(sheetby, i_by, j_by, n_by, sheetproduct, i_merge, j_merge, 0)
    For j = j_do To j_do + j_do_n - 1
        Call productblankby(sheetby, i_by, j_by, n_by, i_do, j, sheetproduct, i_merge, j_merge)
    Next
End Sub
