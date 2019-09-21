Attribute VB_Name = "importmodule"
'namesheet is a sheet of dealerlist, i_code and j_code are the incipent cell of dealer code
Public Sub importdata(namesheet As Worksheet, i_code As Integer, j_code As Integer)
        Dim i, n As Integer
        Dim file_name As String
        Dim summary As Workbook
        Dim sample As Workbook
        
        Set summary = ThisWorkbook   '����
        i = i_code
        n = 1
        
        Do While Not IsEmpty(namesheet.Cells(i, j_code).Value)
                'generate filename
                file_name = ThisWorkbook.path & "\" & namesheet.Cells(i, j_code).Value
                
                'open it
                Set sample = Workbooks.Open(file_name)    '������
                
                'copy your need
                Sheet4.Cells(1, n + 3).Value = Sheet1.Cells(i, j_code + 1).Value
                Call mycopy.copycolumn_n(sample.Sheets("��泵"), summary.Sheets("������תվ"), 59, 6, 2, 3 + n, 3)  '��泵 OK,NG,OK����
                Call mycopy.copycolumn_n(sample.Sheets("PDI"), summary.Sheets("������תվ"), 45, 6, 11, 3 + n, 3)  'PDI OK,NG,OK����
                Call mycopy.copycolumn_n(sample.Sheets("��泵"), summary.Sheets("������תվ"), 8, 6, 18, 3 + n, 50)  '��泵 ����
                Call mycopy.copycolumn_n(sample.Sheets("PDI"), summary.Sheets("������תվ"), 8, 6, 68, 3 + n, 4)  'PDI ����
                summary.Sheets("������תվ").Cells(72, 3 + n).Value = sample.Sheets("PDI").Cells(12, 6).Value
                Call mycopy.copycolumn_n(sample.Sheets("PDI"), summary.Sheets("������תվ"), 21, 6, 73, 3 + n, 23)  'PDI ����
                
                'close it
                sample.Close savechanges:=False
                i = i + 1
                n = n + 1
        Loop
        
        Set summary = Nothing
        Set sample = Nothing
        
        MsgBox ("�������")
        
End Sub
