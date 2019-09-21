VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Shape_check 
   Caption         =   "Shape查看器"
   ClientHeight    =   4395
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   8925
   OleObjectBlob   =   "Shape_check.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Shape_check"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public path As String

Private Sub ComboBox1_Change()
Dim i As Integer
i = ComboBox1.text
ComboBox2.Clear
For i_shape = 1 To ActivePresentation.Slides(i).Shapes.Count
    ComboBox2.AddItem (i_shape)
Next
ActivePresentation.Slides(i).Select
End Sub

Private Sub CommandButton1_Click()
If ComboBox1.text = "" Or ComboBox2.text = "" Then
    MsgBox ("请选择Slide和Shape")
Else
    Call whichshape(ComboBox1.text, ComboBox2.text)
End If
End Sub

Private Sub import_Click()

Dim xlApp As Excel.Application
Set xlApp = New Excel.Application
xlApp.Workbooks.Open path

Dim databook As Excel.Workbook
Set databook = xlApp.Workbooks(1)

Call change.slide1(databook)

xlApp.Workbooks.Close
Set xlApp = Nothing
MsgBox ("finish")

End Sub

Private Sub select_excel_file_Click()
path = openfiledialog.SelectFile
'path = ActivePresentation.path & "\" & "data.xlsx"
End Sub

Private Sub UserForm_Activate()
For i_slide = 1 To ActivePresentation.Slides.Count
    ComboBox1.AddItem (i_slide)
Next
End Sub

Public Sub whichshape(i_slide As Integer, i_shape As Integer)
With ActivePresentation.Slides(i_slide).Shapes(i_shape)
    .Select
End With
End Sub
