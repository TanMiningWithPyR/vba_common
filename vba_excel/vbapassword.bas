Attribute VB_Name = "vbapassword"
Option Explicit

Private Sub vbapassword() '��Ҫ�Ᵽ����Excel�ļ�·��
Dim Filename As String
Filename = Application.GetOpenFilename("Excel�ļ���*.xls & *.xla & *.xlt��,*.xls;*.xla;*.xlt", , "VBA�ƽ�")
If Dir(Filename) = "" Then
MsgBox "û�ҵ�����ļ�,���������á�"
Exit Sub
Else
FileCopy Filename, Filename & ".bak" '�����ļ���
End If
Dim GetData As String * 5
Open Filename For Binary As #1
Dim CMGs As Long
Dim DPBo As Long
Dim i As Long
For i = 1 To LOF(1)
Get #1, i, GetData
If GetData = "CMG=""" Then CMGs = i
If GetData = "[Host" Then DPBo = i - 2: Exit For
Next
If CMGs = 0 Then
MsgBox "���ȶ�VBA��������һ����������...", 32, "��ʾ"
Exit Sub
End If

Dim St As String * 2
Dim s20 As String * 1
'ȡ��һ��0D0Aʮ�������ִ�
Get #1, CMGs - 2, St
'ȡ��һ��20ʮ�����ִ�
Get #1, DPBo + 16, s20
'�滻���ܲ��ݻ���
For i = CMGs To DPBo Step 2
Put #1, i, St
Next
'���벻��Է���
If (DPBo - CMGs) Mod 2 <> 0 Then
Put #1, DPBo + 1, s20
End If
MsgBox "�ļ����ܳɹ�......", 32, "��ʾ"
Close #1
End Sub