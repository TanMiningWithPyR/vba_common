Attribute VB_Name = "Change"
Public Sub slide1(ppt As Presentation, databook As Excel.Workbook)

End Sub
Public Sub slide7(ppt As Presentation, databook As Excel.Workbook)
Dim excelSheet As Excel.Worksheet
Set excelSheet = databook.sheets("计分表")
With ppt.Slides(7).Shapes(6)
    Call mycopy.copyerea_n(excelSheet, .Table, 17, 5, 3, 2, 5, 5)
End With
End Sub
Public Sub slide8(ppt As Presentation, databook As Excel.Workbook)
Dim excelSheet As Excel.Worksheet
Set excelSheet = databook.sheets("运营质量检查打分标准")
With ppt.Slides(8).Shapes(3)
    Call mycopy.copyerea_n(excelSheet, .Table, 4, 12, 2, 4, 14, 4)
    Call mycopy.copyerea_n(excelSheet, .Table, 4, 18, 2, 8, 14, 4)
End With
End Sub
Public Sub slide9(ppt As Presentation, databook As Excel.Workbook)
Dim excelSheet As Excel.Worksheet
Set excelSheet = databook.sheets("运营质量检查打分标准")
With ppt.Slides(9).Shapes(3)
    Call mycopy.copyerea_n(excelSheet, .Table, 18, 12, 2, 4, 12, 4)
    Call mycopy.copyerea_n(excelSheet, .Table, 18, 18, 2, 8, 12, 4)
End With
End Sub
Public Sub slide10(ppt As Presentation, databook As Excel.Workbook)
Dim excelSheet As Excel.Worksheet
Set excelSheet = databook.sheets("运营质量检查打分标准")
With ppt.Slides(10).Shapes(3)
    Call mycopy.copyerea_n(excelSheet, .Table, 30, 12, 2, 4, 10, 4)
    Call mycopy.copyerea_n(excelSheet, .Table, 30, 18, 2, 8, 10, 4)
End With
End Sub
Public Sub slide11(ppt As Presentation, databook As Excel.Workbook)
Dim excelSheet As Excel.Worksheet
Set excelSheet = databook.sheets("运营质量检查打分标准")
With ppt.Slides(11).Shapes(3)
    Call mycopy.copyerea_n(excelSheet, .Table, 40, 12, 2, 4, 10, 4)
    Call mycopy.copyerea_n(excelSheet, .Table, 40, 18, 2, 8, 10, 4)
End With
End Sub
Public Sub slide12(ppt As Presentation, databook As Excel.Workbook)
Dim excelSheet As Excel.Worksheet
Set excelSheet = databook.sheets("运营质量检查打分标准")
With ppt.Slides(12).Shapes(3)
    Call mycopy.copyerea_n(excelSheet, .Table, 50, 12, 2, 4, 8, 4)
    Call mycopy.copyerea_n(excelSheet, .Table, 50, 18, 2, 8, 8, 4)
End With
End Sub
Public Sub slide13(ppt As Presentation, databook As Excel.Workbook)
Dim excelSheet As Excel.Worksheet
Set excelSheet = databook.sheets("运营质量检查打分标准")
With ppt.Slides(13).Shapes(3)
    Call mycopy.copyerea_n(excelSheet, .Table, 58, 12, 2, 4, 10, 4)
    Call mycopy.copyerea_n(excelSheet, .Table, 58, 18, 2, 8, 10, 4)
End With
End Sub
Public Sub slide14(ppt As Presentation, databook As Excel.Workbook)
Dim excelSheet As Excel.Worksheet
Set excelSheet = databook.sheets("运营质量检查打分标准")
With ppt.Slides(14).Shapes(3)
    Call mycopy.copyerea_n(excelSheet, .Table, 68, 12, 2, 4, 4, 4)
    Call mycopy.copyerea_n(excelSheet, .Table, 68, 18, 2, 8, 4, 4)
End With
End Sub
Public Sub slide16plus(ppt As Presentation, databook As Excel.Workbook)
Dim excelSheet As Excel.Worksheet
Dim i_slide As Integer
Dim i_sheet As Integer
Dim text_temple As String
Dim itemNo As String
Dim itemName As String
Dim reason As String

Set excelSheet = databook.sheets("不满足清单")
i_sheet = 6
Do While (Not IsEmpty(excelSheet.Cells(i_sheet, 2).Value)) And (excelSheet.Cells(i_sheet, 2).Value <> "")
    itemNo = excelSheet.Cells(i_sheet, 2).Value
    itemName = excelSheet.Cells(i_sheet, 3).Value
    reason = excelSheet.Cells(i_sheet, 12).Value
    
    With ppt
        .Slides(16 + i_sheet - 6).Copy
        .Slides.Paste (16 + i_sheet - 6 + 1)
        Call changeText(.Slides(16 + i_sheet - 6), itemNo, itemName, reason)
    End With
    
    i_sheet = i_sheet + 1
Loop
ppt.Slides(16 + i_sheet - 6).Delete
End Sub
Private Sub changeText(slide As slide, itemNo As String, itemName As String, reason As String)
With slide.Shapes(2)
    text_temple = .TextFrame.TextRange.Text
    text_temple = Replace(text_temple, "X.X.X", itemNo)
    text_temple = Replace(text_temple, "三级指标", itemName)
    .TextFrame.TextRange.Text = text_temple
End With
With slide.Shapes(3)
    text_temple = .TextFrame.TextRange.Text
    text_temple = Replace(text_temple, "失分原因", reason)
    .TextFrame.TextRange.Text = text_temple
End With
End Sub
