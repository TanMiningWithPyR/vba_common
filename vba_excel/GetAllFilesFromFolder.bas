Attribute VB_Name = "getallfilesfromfolder"
'mypath is 路径,filetype is 类型 (如.png)
Public Sub getall(mypath As String, filetype As String)
 Dim myname$
 Dim filenames()  As String
 Dim i As Integer
 i = 2
  myname = Dir(mypath & "*" & filetype)
  Do While myname <> ""
        ActiveSheet.Cells(i, 1).Value = myname
        i = i + 1
        myname = Dir
  Loop
End Sub
