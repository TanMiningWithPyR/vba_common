Attribute VB_Name = "wordReplace"
Dim xlApp As Object     'Word.Application'
Dim xlDocument As Object    'Word.Document'
Dim xlDocumentCopy As Object
Dim strPath As String    'Project Path
Private Sub Worksheet_SelectionChange(ByVal Target As Range)
strPath = ThisWorkbook.Path
Set xlApp = CreateObject("Word.Application")
For i = 3 To 183
    Set xlDocument = xlApp.Documents.Open(strPath & "\TUV检查通知书模板.docx") '打开模板doc'
    xlDocument.SaveAs strPath & "TUV检查通知书-" & Sheet1.Cells(i, 5).Value & ".docx"
    Set xlDocumentCopy = xlApp.Documents.Open(strPath & "TUV检查通知书-" & Sheet1.Cells(i, 5).Value & ".docx")
    Call replace(Sheet1.Cells(i, 4).Value, Sheet1.Cells(i, 5).Value, Sheet1.Cells(i, 9).Value, Sheet1.Cells(i, 7).Value, Sheet1.Cells(i, 8).Value)
    xlDocumentCopy.Save
    xlDocumentCopy.Close
    Set xlDocumentCopy = Nothing
Next i
End Sub
Private Sub replace(strDealerCode As String, strDealerName As String, strAuditor As String, strStartDate As String, strEndDate As String)
    Set find_range = xlDocument.Content
    find_range.find.Execute findText:="经销商代码", Forward:=True
    If find_range.find.found = True Then
        find_range.Text = strDealerCode
    End If

    Set find_range = xlDocument.Content
    find_range.find.Execute findText:="经销商名称", Forward:=True
    If find_range.find.found = True Then
        find_range.Text = strDealerName
    End If

    Set find_range = xlDocument.Content
    find_range.find.Execute findText:="*C*", Forward:=True
    If find_range.find.found = True Then
        find_range.Text = strAuditor
    End If
    
    Set find_range = xlDocument.Content
    find_range.find.Execute findText:="*D*", Forward:=True
    If find_range.find.found = True Then
        find_range.Text = strStartDate
    End If
    
    Set find_range = xlDocument.Content
    find_range.find.Execute findText:="*E*", Forward:=True
    If find_range.find.found = True Then
        find_range.Text = strEndDate
    End If
End Sub


