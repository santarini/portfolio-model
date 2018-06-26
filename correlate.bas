Sub MonthCorrelate()

Dim baseData, corrData, topCell As Range
Dim countx, county As Integer
Dim CorrelationVar As Double
Dim str As String

Sheets.Add.Name = "MonthlyCorr"
Set topCell = Range("A1")
countx = 1
county = 1

For Each Basesheet In Worksheets
    If InStr(1, Basesheet.Name, "(Mon)") > 0 Then
        
        Basesheet.Activate
        
        'find "Sum of Intraday"
        
            Cells.Find(What:="Sum of Intraday", After:=ActiveCell, LookIn:=xlFormulas _
                , LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
                MatchCase:=False, SearchFormat:=False).Activate
        
        'select data beneath
            
            Selection.Offset(1, 0).Select
            Range(Selection, Selection.End(xlDown)).Select
            
        'selection equals usableData
            Set baseData = Selection

        'get average and stddev of selection
        
            BaseArithmeticMean = Application.WorksheetFunction.Average(baseData)
            BaseStandardDeviation = Application.WorksheetFunction.StDev_P(baseData)
        
        'paste stats into monthsummary
            
            Worksheets("MonthlyCorr").Select
            str = Split(Basesheet.Name, "(")(0) & vbNewLine & Chr(181) & "=" & Format(BaseArithmeticMean, "Percent") & " " & ChrW(&H3C3) & "=" & Format(BaseStandardDeviation, "Percent")
            topCell.Offset(0, countx).WrapText = True
            topCell.Offset(0, countx).Value = str
            

                
                For Each CurrentSheet In Worksheets
                    If InStr(1, CurrentSheet.Name, "(Mon)") > 0 Then
                            CurrentSheet.Activate
    
                            'find "Sum of Intraday"
                            
                                Cells.Find(What:="Sum of Intraday", After:=ActiveCell, LookIn:=xlFormulas _
                                    , LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
                                    MatchCase:=False, SearchFormat:=False).Activate
                            
                            'select data beneath
                                
                                Selection.Offset(1, 0).Select
                                Range(Selection, Selection.End(xlDown)).Select
                                
                            'selection equals usableData
                                Set corrData = Selection
                                
                            'get average and stddev of selection
        
                                TrgtArithmeticMean = Application.WorksheetFunction.Average(corrData)
                                TrgtStandardDeviation = Application.WorksheetFunction.StDev_P(corrData)
                                
                            'find correlation
                                CorrelationVar = Application.WorksheetFunction.Correl(baseData, corrData)
                                
                            'navigate to "MonthlyCorr"
                                Worksheets("MonthlyCorr").Select
                                
                            'paste corrData name in row
                                str = Split(CurrentSheet.Name, "(")(0) & vbNewLine & Chr(181) & "=" & Format(TrgtArithmeticMean, "Percent") & " " & ChrW(&H3C3) & "=" & Format(TrgtStandardDeviation, "Percent")
                                topCell.Offset(county, 0).Value = str
                                
                            'paste correlation
                                topCell.Offset(county, countx).Value = CorrelationVar
                                county = county + 1
                    
                    End If
                Next
            county = 1
            countx = countx + 1
        Basesheet.Activate
    End If
Next
Worksheets("MonthlyCorr").Activate

'center first column and row
Rows("1:1").Select
    With Selection
        .HorizontalAlignment = xlCenter
    End With
Columns("A:A").Select
    With Selection
        .HorizontalAlignment = xlCenter
    End With
    
'format numbers and cells
Range("B2").Select
Range(Selection, Selection.End(xlDown)).Select
Range(Selection, Selection.End(xlToRight)).Select
Selection.NumberFormat = "0.00"
With Selection
    .HorizontalAlignment = xlCenter
    .VerticalAlignment = xlCenter
End With

'heat map
    Selection.FormatConditions.AddColorScale ColorScaleType:=3
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    Selection.FormatConditions(1).ColorScaleCriteria(1).Type = _
        xlConditionValueNumber
    Selection.FormatConditions(1).ColorScaleCriteria(1).Value = -1
    With Selection.FormatConditions(1).ColorScaleCriteria(1).FormatColor
        .Color = 7039480
        .TintAndShade = 0
    End With
    Selection.FormatConditions(1).ColorScaleCriteria(2).Type = _
        xlConditionValueNumber
    Selection.FormatConditions(1).ColorScaleCriteria(2).Value = 0
    With Selection.FormatConditions(1).ColorScaleCriteria(2).FormatColor
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = 0
    End With
    Selection.FormatConditions(1).ColorScaleCriteria(3).Type = _
        xlConditionValueNumber
    Selection.FormatConditions(1).ColorScaleCriteria(3).Value = 1
    With Selection.FormatConditions(1).ColorScaleCriteria(3).FormatColor
        .Color = 8109667
        .TintAndShade = 0
    End With

'autofit rows and cols
Cells.Select
Cells.EntireColumn.AutoFit
Cells.EntireRow.AutoFit

'reset selectio

Range("B1").Select
Range(Selection, Selection.End(xlToRight)).Select
Set HeaderRng = Selection

For Each cell In HeaderRng
openPos = InStr(cell, "=")
closePos = InStr(cell, "%")
AssetReturn = Mid(cell, openPos + 1, closePos - openPos - 1)
If AssetReturn <= 0 Then
    cell.Characters(openPos + 1, closePos - openPos - 1).Font.Color = RGB(255, 0, 0)
End If

If AssetReturn >= 0 Then
    cell.Characters(openPos + 1, closePos - openPos - 1).Font.Color = RGB(0, 190, 0)
End If
Next


Range("A2").Select
Range(Selection, Selection.End(xlDown)).Select
Set HeaderRng = Selection

For Each cell In HeaderRng
openPos = InStr(cell, "=")
closePos = InStr(cell, "%")
AssetReturn = Mid(cell, openPos + 1, closePos - openPos - 1)
If AssetReturn <= 0 Then
    cell.Characters(openPos + 1, closePos - openPos - 1).Font.Color = RGB(255, 0, 0)
End If

If AssetReturn >= 0 Then
    cell.Characters(openPos + 1, closePos - openPos - 1).Font.Color = RGB(0, 190, 0)
End If
Next

Range("A2").Select


End Sub
