Sub makeMonths()
Dim CurrentSheet As Worksheet
Dim SheetName As String
For Each CurrentSheet In Worksheets
    If InStr(1, CurrentSheet.Name, "(D)") > 0 Then
        CurrentSheet.Activate
        SheetName = Split(CurrentSheet.Name, "(")(0)
        Call monthlyData(SheetName)
        CurrentSheet.Activate
    End If
Next
End Sub
Function monthlyData(SheetName As String)

Dim PTTop, PTYrs, PTMons, cell, CellWorking As Range
Dim PTRowCount As Integer
Dim WBTB As Worksheet
Dim str As String
Dim useableData As Range

Range("A1").Select

'Insert pivot table
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Set useableData = Selection

    Range("A1").Select
    
    Sheets.Add.Name = SheetName & "(Mon)"
    
    
    Sheets(SheetName & "(Mon)").Activate
        ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:= _
        useableData, Version:=6).CreatePivotTable TableDestination:= _
        Sheets(SheetName & "(Mon)").Range("A3"), TableName:="PivotTable1", DefaultVersion:=6
    Sheets(SheetName & "(Mon)").Select
    Cells(3, 1).Select
    
'Configure pivot table
    With ActiveSheet.PivotTables("PivotTable1").PivotFields("date")
        .Orientation = xlRowField
        .Position = 1
    End With
    ActiveSheet.PivotTables("PivotTable1").PivotFields("date").AutoGroup
    ActiveSheet.PivotTables("PivotTable1").AddDataField ActiveSheet.PivotTables( _
        "PivotTable1").PivotFields("Intraday Open to Close Percent"), "Sum of Intraday %", xlSum
    Cells(4, 1).Select
    Selection.Group Start:=True, End:=True, Periods:=Array(False, False, False, _
        False, True, False, True)
        
'Find "Row Labels"
    Cells.Find(What:="Row Labels", After:=ActiveCell, LookIn:=xlFormulas, _
        LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
        MatchCase:=False, SearchFormat:=False).Activate

'set cell to Rng

    Set PTTop = Selection
    PTTop.Value = "Month"
    

'select content to bottom
    Range(Selection, Selection.End(xlDown)).Select

'RowCount
    Set PTMons = Selection

'select content to right
    Range(Selection, Selection.End(xlToRight)).Select

'copy selection

    Selection.Copy

'paste values

    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False

'insert a column to the left push content right

    Columns("A:A").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove

'Select PTMons
    PTMons.Activate
    PTMons.Offset(0, -1).Select
    Set PTYrs = Selection
    
'Add Label to top of year column
    PTTop.Offset(0, -1).Value = "Year"

'Rng activate
    PTTop.Select

For Each cell In PTMons
    If IsNumeric(cell) Then
        cell.Select
        Selection.Cut
        cell.Offset(0, -1).Select
        ActiveSheet.Paste
    End If
Next


For Each cell In PTYrs
    If IsEmpty(cell) Then
        cell.Value = cell.Offset(-1, 0).Value
    End If
Next


For Each cell In PTMons
    If IsEmpty(cell) Or InStr(cell, "Grand Total") > 0 Then
        cell.Select
        cell.EntireRow.Delete
    End If
Next

'delete the top rows
    Rows("1:2").Delete

'format the data
    
    Columns("C:C").Select
    Selection.Style = "Percent"
    Selection.NumberFormat = "0.000%"
    
're-center A1
    Range("A1").Select
    

End Function
Function monthlySummary(SheetName As String)

    Dim useableData As Range
    Dim MonthlyArithmeticMean, MonthlyStandardDeviation, MonthlyVariance As Double
    
    
'find "Sum of Intraday"

    Cells.Find(What:="Sum of Intraday", After:=ActiveCell, LookIn:=xlFormulas _
        , LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
        MatchCase:=False, SearchFormat:=False).Activate

'select data beneath
    
    Selection.Offset(1, 0).Select
    Range(Selection, Selection.End(xlDown)).Select
    
'selection equals usableData
    Set useableData = Selection

    MonthlyArithmeticMean = Application.WorksheetFunction.Average(useableData)
    MonthlyStandardDeviation = Application.WorksheetFunction.StDev_P(useableData)
    MonthlyVariance = MonthlyStandardDeviation * MonthlyStandardDeviation

'worksheet dailySummary select
    Worksheets("MonthSummary").Activate
    Range("A1").Select
    If IsEmpty(Range("A1").Offset(0, 1)) Then
        Range("B1").Select
    Else
        Selection.End(xlToRight).Select
        Selection.Offset(0, 1).Select
    End If
    
    ActiveCell.Value = SheetName
    
    ActiveCell.Offset(1, 0).Value = MonthlyArithmeticMean
    ActiveCell.Offset(2, 0).Value = MonthlyStandardDeviation
    ActiveCell.Offset(3, 0).Value = MonthlyVariance

End Function
