Attribute VB_Name = "mytable"
Option Explicit

    Public tablename As String
    Private tablerange As Range
    Private i_rowname As Integer
    Private j_columnname As Integer
    Private ColofRowAttribute As Integer
    Private RowofColAttribute As Integer
    Private sheettable As Worksheet
       
    'row_i and column_i is the incipient of the table index, row_e and columnname_e is the end of the table index
    
    Public Sub table_initalize(sheet As Worksheet, ByVal row_i As Integer, ByVal column_i As Integer, ByVal row_e As Integer, ByVal column_e As Integer, ByVal get_Row_attr_InWhichCol As Integer, ByVal get_Col_attr_InWhichRow As Integer)
          Dim cell_i, cell_e As Range
          Dim Strcell_i, Strcell_e As String
          
          i_rowname = row_i
          j_columnname = column_i
          
          Set cell_i = sheet.Cells(row_i, column_i)
          Set cell_e = sheet.Cells(row_e, column_e)
          Strcell_i = cell_i.Address()
          Strcell_e = cell_e.Address()
          Set sheettable = sheet
          Set tablerange = sheet.Range(Strcell_i & ":" & Strcell_e)
          
          ColofRowAttribute = get_Row_attr_InWhichCol '确定行属性是那一列
          RowofColAttribute = get_Col_attr_InWhichRow '确定列属性是那一行
          
    End Sub
   Private Function gettablecellindex_i(i As String) As Integer
   
        gettablecellindex_i = findij.IntFind_str_row(i, sheettable, i_rowname, ColofRowAttribute)
          
   End Function
   Private Function gettablecellindex_j(j As String) As Integer
        
        gettablecellindex_j = findij.IntFind_str_col(j, sheettable, RowofColAttribute, j_columnname)

   End Function
   Public Function mycells(rowname As String, columnname As String) As Variant
        Dim i, j As Integer
        i = gettablecellindex_i(rowname)
        j = gettablecellindex_j(columnname)
        mycells = sheettable.Cells(i, j).Value2
        
   End Function
