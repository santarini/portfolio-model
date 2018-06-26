Function CapitalAllocationLine()

Range("A1").Select
Selection.End(xlDown).Select
Set OptimalRng = Selection
OptimalReturn = OptimalRng.Offset(0, 3)
OptimalStDev = OptimalRng.Offset(0, 4)


Cells.Find(What:="Individual Stats", After:=ActiveCell, LookIn:= _
    xlFormulas, LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:= _
    xlNext, MatchCase:=False, SearchFormat:=False).Activate
Selection.End(xlDown).Select
riskFreeRate = Selection.Offset(0, 1)

Range("I1").Value = "Risky Weight"
Range("J1").Value = "RFR Weight"

j = 1
k = 0
For i = 1 To 17
    Range("I1").Offset(i, 0).Value = k
    Range("J1").Offset(i, 0).Value = j
    PortfolioReturn = ((k * OptimalReturn) + (j * riskFreeRate))
    PortfolioStdDev = (k * OptimalStDev)
    Range("K1").Offset(i, 0).Value = PortfolioReturn
    Range("L1").Offset(i, 0).Value = PortfolioStdDev
    j = j - 0.1
    k = k + 0.1
Next

Range("I2:J2").Select
Range(Selection, Selection.End(xlDown)).Select

Selection.NumberFormat = "0%"

Range("K1").Value = "Return Portfolio"
Range("L1").Value = "Portfolio StDev"
    ActiveSheet.ChartObjects("Chart 1").Activate
    ActiveChart.ChartArea.Select
    Application.CutCopyMode = False
    Application.CutCopyMode = False
    ActiveChart.SeriesCollection.NewSeries
    ActiveChart.FullSeriesCollection(4).XValues = "=Portfolio!$L$2:$L$18"
    ActiveChart.FullSeriesCollection(4).Values = "=Portfolio!$K$2:$K$18"
    ActiveChart.FullSeriesCollection(4).Select
    Selection.MarkerStyle = -4142
    ActiveChart.ChartArea.Select

Range("A1").Select



End Function

