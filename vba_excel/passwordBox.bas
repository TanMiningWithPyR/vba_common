Attribute VB_Name = "passwordBox"
Public Sub inputBox_password(sheetname As String)
Sheets(sheetname).Range("A:AZ").EntireColumn.Hidden = True
If Application.InputBox("���������Ȩ������:") = "hongmenyan" Then
Range("A:AZ").EntireColumn.Hidden = False
Else
MsgBox "�������,�����˳�!"
Sheets("���ڻ��ܱ�").Select
End If
End Sub

