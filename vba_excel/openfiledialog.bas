Attribute VB_Name = "openfiledialog"
Public Function SelectFile() As String
    'ѡ��һ�ļ�
    With Application.FileDialog(msoFileDialogFilePicker)
        .AllowMultiSelect = False
        '��ѡ��
        .Filters.Clear
        '����ļ�������
        .Filters.Add "Excel Files", "*.xls;*.xlsm"
        .Filters.Add "All Files", "*.*"
        '���������ļ�������
        If .Show = -1 Then
            'FileDialog ����� Show ������ʾ�Ի��򣬲��ҷ��� -1��������� OK���� 0��������� Cancel����
            MsgBox "��ѡ����ļ��ǣ�" & .SelectedItems(1), vbOKOnly + vbInformation, "����Excel"
            SelectFile = .SelectedItems(1)
        End If
    End With
End Function
Public Sub SelectFiles()
    'ѡ�����ļ�
    'www.okexcel.com.cn
    Dim l As Long
    With Application.FileDialog(msoFileDialogFilePicker)
        .AllowMultiSelect = True
        '��ѡ��
        .Filters.Clear
        '����ļ�������
        .Filters.Add "Excel Files", "*.xls;*.xlsm"
        .Filters.Add "All Files", "*.*"
        '���������ļ�������
        .Show
        'FileDialog ����� Show ������ʾ�Ի��򣬲��ҷ��� -1��������� OK���� 0��������� Cancel����
        For l = 1 To .SelectedItems.Count
            MsgBox "��ѡ����ļ��ǣ�" & .SelectedItems(l), vbOKOnly + vbInformation, "����Excel"
        Next
    End With
End Sub
Public Function SelectFolder() As String
    'ѡ��һ�ļ�
    'www.okexcel.com.cn
    With Application.FileDialog(msoFileDialogFolderPicker)
        If .Show = -1 Then
        'FileDialog ����� Show ������ʾ�Ի��򣬲��ҷ��� -1��������� OK���� 0��������� Cancel����
            MsgBox "��ѡ����ļ����ǣ�" & .SelectedItems(1), vbOKOnly + vbInformation, "����Excel"
            SelectFolder = .SelectedItems(1)
        End If
    End With
End Function
