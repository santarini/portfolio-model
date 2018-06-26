Function optimalPortfolio()

Dim riskFreeRate As Double
Dim OptimalRng As Range


riskFreeRate = InputBox("What is the Annual Risk Free Rate?", "Risk Free Rate of Return", 1)
riskFreeRate = riskFreeRate / 12

Range("F7").Value = "Risk Free Rate"
Range("G7").Value = riskFreeRate / 100

CovAB = Range("G5").Value
CorrAB = Range("G6").Value

AReturn = Range("G2").Value
AVar = Range("G3").Value
AStDev = Range("G4").Value

BReturn = Range("H2").Value
BVar = Range("H3").Value
BStDev = Range("H4").Value

AOptimalW = (((AReturn - riskFreeRate) * BVar) - ((BReturn - riskFreeRate) * CovAB)) / ((((AReturn - riskFreeRate) * BVar) + ((BReturn - riskFreeRate) * AVar)) - ((AReturn - riskFreeRate + BReturn - riskFreeRate) * CovAB))
BOptimalW = 1 - AOptimalW

PortfolioReturn = ((AOptimalW * AReturn) + (BOptimalW * BReturn))
PortfolioStdDev = Sqr(((AOptimalW * AStDev) ^ 2) + ((BOptimalW * BStDev) ^ 2) + (2 * AOptimalW * BOptimalW * CorrAB * AStDev * BStDev))

Range("A1").Select
Selection.End(xlDown).Select
Selection.Offset(1, 0).Select
Set OptimalRng = Selection
OptimalRng.Value = "Optimal"
OptimalRng.Offset(0, 1).Value = AOptimalW
OptimalRng.Offset(0, 2).Value = BOptimalW
OptimalRng.Offset(0, 3).Value = PortfolioReturn
OptimalRng.Offset(0, 4).Value = PortfolioStdDev
OptimalRng.Offset(0, 1).Select
Range(Selection, Selection.End(xlToRight)).Select
Selection.NumberFormat = "0.00%"
Selection.Interior.Color = RGB(255, 255, 204)


ActiveSheet.ChartObjects("Chart 1").Activate
ActiveChart.SeriesCollection.NewSeries
ActiveChart.FullSeriesCollection(3).XValues = OptimalRng.Offset(0, 4)
ActiveChart.FullSeriesCollection(3).Values = OptimalRng.Offset(0, 3)
ActiveSheet.ChartObjects("Chart 1").Activate
ActiveChart.FullSeriesCollection(3).Select
ActiveChart.FullSeriesCollection(3).Points(1).Select
ActiveChart.FullSeriesCollection(3).Select
With Selection.Format.Line
    .Visible = msoTrue
    .ForeColor.RGB = RGB(255, 0, 0)
    .Transparency = 0
End With
With Selection.Format.Fill
    .Visible = msoTrue
    .ForeColor.RGB = RGB(255, 0, 0)
    .Transparency = 0
    .Solid
End With

Range("A1").Select

End Function
