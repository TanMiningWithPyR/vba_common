VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "mytable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private table As String          'table name
Private row_attri As String         'row attribute
Private column_attri As String      'column attribute
Private indexofrow_b_i, indexofrow_b_j, indexofcolumn_b_i, indexofcolumn_b_j As Integer   'row index name begin (i,j) and column index name begin (i,j)
Private RowIndexLength, ColumnIndexLength As Integer
Private tablesheet As Worksheet

'table name property
Public Property Get Tablename() As String
    Tablename = table
End Property
Public Property Let Tablename(ByVal T As String)
    table = T
End Property
'row attribute property and column attribute property
Public Property Get Row_attribute() As String
    Row_attribute = row_attri
End Property
Public Property Get Column_attribute() As String
    Column_attribute = column_attri
End Property
Public Property Let Row_attribute(ByVal R As String)
    row_attri = R
End Property
Public Property Let Column_attribute(ByVal R As String)
    column_attri = R
End Property
'row attribute name
Public Property Get Beginingofrow_i() As Integer
    Beginingofrow_i = indexofrow_b_i
End Property
Public Property Let Beginingofrow_i(ByVal i As Integer)
    indexofrow_b_i = i
End Property
Public Property Get Beginingofrow_j() As Integer
    Beginingofrow_j = indexofrow_b_j
End Property
Public Property Let Beginingofrow_j(ByVal j As Integer)
    indexofrow_b_j = j
End Property
'column attribute name
Public Property Get Beginingofcolumn_i() As Integer
    Beginingofcolumn_i = indexofcolumn_b_i
End Property
Public Property Let Beginingofcolumn_i(ByVal i As Integer)
    indexofcolumn_b_i = i
End Property
Public Property Get Beginingofcolumn_j() As Integer
    Beginingofcolumn_j = indexofcolumn_b_j
End Property
Public Property Let Beginingofcolumn_j(ByVal j As Integer)
    indexofcolumn_b_j = j
End Property
'row and column length
Public Property Get Rowlength() As Integer
    Rowlength = RowIndexLength
End Property
Public Property Let Rowlength(ByVal n As Integer)
    RowIndexLength = n
End Property
Public Property Get Columnlength() As Integer
    Columnlength = ColumnIndexLength
End Property
Public Property Let Columnlength(ByVal n As Integer)
    ColumnIndexLength = n
End Property
' which sheet is table in
Public Property Get TableInsheet() As Worksheet
    Set TableInsheet = tablesheet
End Property
Public Property Let TableInsheet(Sheet As Worksheet)
    Set tablesheet = Sheet
End Property
' class initial
Private Sub Class_Initialize()
    Set tablesheet = ActiveSheet
End Sub
' generate tablerange
Public Function Tablerange() As Range
    Dim Alphacode_i, Alphacode_e As String   'such as "A2","D4",table initial and end
    Alphacode_i = tablesheet.Cells(indexofrow_b_i, indexofcolumn_b_j).Address()
    Alphacode_e = tablesheet.Cells(indexofrow_b_i + RowIndexLength, indexofcolumn_b_j + ColumnIndexLength).Address()
    Set Tablerange = tablesheet.Range(Alphacode_i & ":" & Alphacode_e)
End Function
' get value
Public Function mycells(rowname As String, columnname As String) As Variant
    Dim i, j As Integer
    i = findij.IntFind_str_row_n(rowname, tablesheet, Beginingofrow_i, Beginingofrow_j, Rowlength)
    j = findij.IntFind_str_col_n(columnname, tablesheet, Beginingofcolumn_i, Beginingofcolumn_j, Columnlength)
    mycells = tablesheet.Cells(i, j).Value2
End Function

