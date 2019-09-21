Attribute VB_Name = "openfiledialog"
Public Function SelectFile() As String
    '选择单一文件
    With Application.FileDialog(msoFileDialogFilePicker)
        .AllowMultiSelect = False
        '单选择
        .Filters.Clear
        '清除文件过滤器
        .Filters.Add "Excel Files", "*.xls;*.xlsm"
        .Filters.Add "All Files", "*.*"
        '设置两个文件过滤器
        If .Show = -1 Then
            'FileDialog 对象的 Show 方法显示对话框，并且返回 -1（如果您按 OK）和 0（如果您按 Cancel）。
            MsgBox "您选择的文件是：" & .SelectedItems(1), vbOKOnly + vbInformation, "智能Excel"
            SelectFile = .SelectedItems(1)
        End If
    End With
End Function
Public Sub SelectFiles()
    '选择多个文件
    'www.okexcel.com.cn
    Dim l As Long
    With Application.FileDialog(msoFileDialogFilePicker)
        .AllowMultiSelect = True
        '单选择
        .Filters.Clear
        '清除文件过滤器
        .Filters.Add "Excel Files", "*.xls;*.xlsm"
        .Filters.Add "All Files", "*.*"
        '设置两个文件过滤器
        .Show
        'FileDialog 对象的 Show 方法显示对话框，并且返回 -1（如果您按 OK）和 0（如果您按 Cancel）。
        For l = 1 To .SelectedItems.Count
            MsgBox "您选择的文件是：" & .SelectedItems(l), vbOKOnly + vbInformation, "智能Excel"
        Next
    End With
End Sub
Public Function SelectFolder() As String
    '选择单一文件
    'www.okexcel.com.cn
    With Application.FileDialog(msoFileDialogFolderPicker)
        If .Show = -1 Then
        'FileDialog 对象的 Show 方法显示对话框，并且返回 -1（如果您按 OK）和 0（如果您按 Cancel）。
            MsgBox "您选择的文件夹是：" & .SelectedItems(1), vbOKOnly + vbInformation, "智能Excel"
            SelectFolder = .SelectedItems(1)
        End If
    End With
End Function
