Function createPortfolio()

Dim MainRng As Range
Dim AssetA, AssetB As Range
Dim AssetAName, AssetBName As String
Dim AssetAReturn, AssetBReturn, AssetAStD, AssetBStD As Double
Dim openPos, closePos, DataRowCount As Integer


Set MainRng = Selection

'get asset a
Selection.End(xlUp).Select
Set AssetA = Selection


'extract asset a data
AssetAName = Split(AssetA.Value, Chr(181))(0)

openPos = InStr(AssetA, Chr(181))
closePos = InStr(AssetA, "%")
AssetAReturn = Mid(AssetA, openPos + 2, closePos - openPos - 2) / 100

openPos = InStr(1, AssetA, "=")
openPos = InStr(openPos + 1, AssetA, "=")
closePos = InStr(1, AssetA, "%")
closePos = InStr(closePos + 1, AssetA, "%")
AssetAStD = Mid(AssetA, openPos + 1, closePos - openPos - 2) / 100


MainRng.Select

'get asset b
Selection.End(xlToLeft).Select
Set AssetB = Selection

'extract asset b data
AssetBName = Split(AssetB.Value, Chr(181))(0)

openPos = InStr(AssetB, Chr(181))
closePos = InStr(AssetB, "%")
AssetBReturn = Mid(AssetB, openPos + 2, closePos - openPos - 2) / 100

openPos = InStr(1, AssetB, "=")
openPos = InStr(openPos + 1, AssetB, "=")
closePos = InStr(1, AssetB, "%")
closePos = InStr(closePos + 1, AssetB, "%")
AssetBStD = Mid(AssetB, openPos + 1, closePos - openPos - 2) / 100

Sheets.Add.Name = "Portfolio"

Range("A1").Value = AssetAName & " Weight"
Range("B1").Value = AssetBName & " Weight"

j = 1
k = 0
For i = 1 To 11
    Range("A1").Offset(i, 0).Value = k
    Range("B1").Offset(i, 0).Value = j
    j = j - 0.1
    k = k + 0.1
Next

Range("A2:B2").Select
Range(Selection, Selection.End(xlDown)).Select

Selection.NumberFormat = "0%"

Range("C1").Value = "Portfolio Return"
Range("D1").Value = "Portfolio StDev"

For i = 1 To 11
    AWeight = Range("A1").Offset(i, 0).Value
    BWeight = Range("B1").Offset(i, 0).Value
    PortfolioReturn = ((AWeight * AssetAReturn) + (BWeight * AssetBReturn))
    PortfolioStdDev = Sqr(((AWeight * AssetAStD) ^ 2) + ((BWeight * AssetBStD) ^ 2) + (2 * AWeight * BWeight * MainRng.Value * AssetAStD * AssetBStD))
    Range("C1").Offset(i, 0).Value = PortfolioReturn
    Range("D1").Offset(i, 0).Value = PortfolioStdDev
Next

Range("C2:D2").Select
Range(Selection, Selection.End(xlDown)).Select

Selection.NumberFormat = "0%"

DataRowCount = Selection.Rows.Count

Range("E1").Value = "Individual Stats"
Range("E2").Value = "Average Return"
Range("E3").Value = "Variance"
Range("E4").Value = "StDev"
Range("E5").Value = "Cov"
Range("E6").Value = "Corr"

Range("F1").Value = AssetAName
Range("F2").Value = AssetAReturn
Range("F3").Value = (AssetAStD ^ 2)
Range("F4").Value = AssetAStD
Range("F5").Value = MainRng.Value / (AssetAStD * AssetBStD)
Range("F6").Value = MainRng.Value

Range("G1").Value = AssetBName
Range("G2").Value = AssetBReturn
Range("G3").Value = (AssetBStD ^ 2)
Range("G4").Value = AssetBStD
Range("G5").Value = MainRng.Value / (AssetAStD * AssetBStD)
Range("G6").Value = MainRng.Value

'Columns("C:C").Select
'Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
'Range("C1").Value = "Asset Weights"
'For i = 1 To DataRowCount
    'DataLabel = AssetAName & " " & Format(Range("C1").Offset(i, -2).Value, "Percent") & ", " & AssetBName & " " & Format(Range("C1").Offset(i, -1).Value, "Percent")
    'Range("C1").Offset(i, 0).Value = DataLabel
'Next
'Columns("C:C").Select
'With Selection
'    .WrapText = False
'End With
Rows("1:1").Select
With Selection
    .WrapText = False
End With
'Columns("A:B").Delete

Columns("A:A").Select
Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
Range("A1").Value = "ID"
For i = 1 To DataRowCount
    Range("A1").Offset(i, 0).Value = Chr(i + 64)
Next

    ActiveSheet.Shapes.AddChart2(240, xlXYScatterSmooth).Select
    ActiveChart.SeriesCollection.NewSeries
    ActiveChart.FullSeriesCollection(1).XValues = "=Portfolio!$E$2:$E$" & (DataRowCount + 1)
    ActiveChart.FullSeriesCollection(1).Values = "=Portfolio!$D$2:$D$" & (DataRowCount + 1)
    ActiveChart.HasTitle = True
    ActiveChart.ChartTitle.Text = "Efficient Frontier"
    With ActiveChart.Axes(xlValue)
     .HasTitle = True
     With .AxisTitle
     .Caption = "Portfolio " & Chr(181)
     End With
    End With
    With ActiveChart.Axes(xlCategory)
     .HasTitle = True
     .AxisTitle.Caption = "Portfolio " & ChrW(&H3C3)
    End With
    ActiveChart.Axes(xlValue).MajorUnit = 0.01
    ActiveChart.Axes(xlValue).MinorUnit = 0.005
    ActiveChart.FullSeriesCollection(1).Select
    ActiveChart.SetElement (msoElementDataLabelTop)
    ActiveSheet.ChartObjects("Chart 1").Activate

    ActiveChart.FullSeriesCollection(1).DataLabels.Select
    ActiveChart.SeriesCollection(1).DataLabels.Format.TextFrame2.TextRange. _
        InsertChartField msoChartFieldRange, "=Portfolio!$A$2:$A$" & (DataRowCount + 1), 0
    Selection.ShowRange = True
    Selection.ShowValue = False
    
Cells.Select
Selection.Columns.AutoFit
Selection.Rows.AutoFit

Range("A1").Select
    
End Function

