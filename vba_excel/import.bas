Attribute VB_Name = "import"
'namesheet is a dealerlist, i_code and j_code are the incipent cell of dealer code
Public Sub importdata(namesheet As Worksheet, i_code As Integer, j_code As Integer)
        Dim i, n As Integer
        Dim file_name As String
        i = i_code
        n = 1
        Do While Not IsEmpty(namesheet.Cells(i, j_code).Value)
                'generate filename
                file_name = "D:\c\Desktop\system\宾利\Bentley现场版检查表（更新）\宾利" & namesheet.Cells(i, 6).Value & ".xlsx"
                'open it
                Workbooks.Open (file_name)
                'copy your need
                Call mycopy.copycolumn_n(Workbooks(2).Sheets(2), Workbooks(1).Sheets(2), 1, 12, 1, 4 + n, 174)
                Workbooks(1).Sheets(2).Cells(2, 4 + n).Value = Workbooks(1).Sheets(1).Cells(i, j_code).Value
                'close it
                Workbooks(2).Close savechanges:=False
                i = i + 1
                n = n + 1
        Loop
End Sub
