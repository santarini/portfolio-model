Sub scatterPlot()

Dim ColCount, i As Integer
Dim SelectData As Range
Dim mean, sd As Double
Dim Chart1 As Chart
Dim name As String

Range("A1").Select
Range(Selection, Selection.End(xlDown)).Select
Range(Selection, Selection.End(xlToRight)).Select
ColCount = Selection.Columns.Count
Range("A6").Select

ActiveSheet.Shapes.AddChart2(240, xlXYScatter).Select


For i = 1 To ColCount - 1:
    name = Range("A1").Offset(0, i)
    mean = Range("A1").Offset(1, i)
    sd = Range("A1").Offset(2, i)
    ActiveChart.SeriesCollection.NewSeries
    ActiveChart.FullSeriesCollection(i).name = name
    ActiveChart.FullSeriesCollection(i).XValues = sd
    ActiveChart.FullSeriesCollection(i).Values = mean
    ActiveChart.FullSeriesCollection(i).Select
    ActiveChart.SetElement (msoElementDataLabelLeft)
    ActiveChart.FullSeriesCollection(i).DataLabels.Select
    Selection.ShowSeriesName = True
    Selection.ShowValue = False
Next i
ActiveChart.HasTitle = True
ActiveChart.ChartTitle.Text = "Average Monthly Return by StDev"
With ActiveChart.Axes(xlValue)
    .HasTitle = True
    With .AxisTitle
        .Caption = Chr(181)
    End With
End With
With ActiveChart.Axes(xlCategory)
    .HasTitle = True
    .AxisTitle.Caption = ChrW(&H3C3)
End With
ActiveChart.Axes(xlValue).Select
Selection.TickLabels.NumberFormat = "0.00%"

ActiveChart.Axes(xlCategory).Select
Selection.TickLabels.NumberFormat = "0.00%"

End Sub
