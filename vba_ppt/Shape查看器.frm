VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Shape�鿴�� 
   Caption         =   "Shape�鿴��"
   ClientHeight    =   3030
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4560
   OleObjectBlob   =   "Shape�鿴��.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Shape�鿴��"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub ComboBox1_Change()
Dim i As Integer
i = ComboBox1.Text
ComboBox2.Clear
For i_shape = 1 To ActivePresentation.Slides(i).Shapes.Count
    ComboBox2.AddItem (i_shape)
Next
ActivePresentation.Slides(i).Select
End Sub

Private Sub CommandButton1_Click()
If ComboBox1.Text = "" Or ComboBox2.Text = "" Then
    MsgBox ("��ѡ��Slide��Shape")
Else
    Call checkshape.whichshape(ComboBox1.Text, ComboBox2.Text)
End If
End Sub

Private Sub UserForm_Activate()
For i_slide = 1 To ActivePresentation.Slides.Count
    ComboBox1.AddItem (i_slide)
Next
End Sub

