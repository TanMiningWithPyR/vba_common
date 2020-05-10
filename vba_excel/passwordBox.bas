Attribute VB_Name = "passwordBox"
Public Sub inputBox_password(sheetname As String)
Sheets(sheetname).Range("A:AZ").EntireColumn.Hidden = True
If Application.InputBox("请输入操作权限密码:") = "hongmenyan" Then
Range("A:AZ").EntireColumn.Hidden = False
Else
MsgBox "密码错误,即将退出!"
Sheets("排期汇总表").Select
End If
End Sub

