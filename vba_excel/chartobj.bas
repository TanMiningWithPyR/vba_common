Attribute VB_Name = "chartobj"
Option Explicit

Public Sub changechartsource(chartobj As ChartObject, chartrange As Range)
    chartobj.Chart.SetSourceData Source:=chartrange
End Sub

