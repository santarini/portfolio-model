Sub compileByPath()

Dim FolderPath As String
Dim PathCountCondition As String
Dim FileName As String
Dim count As Integer
Dim FileNumber As Integer
Dim MainWB As Workbook
Dim WB As Workbook
Dim Rng As Range
Dim RngNoPath As String
Dim StartTime As Double
Dim SecondsElapsed As Double
Dim tickersPerSec As Double
Dim SummaryRng As Range
Dim CurrentSheet As Worksheet
Dim SheetName As String


StartTime = Timer

'set this workbook as the main workbook

Set MainWB = ActiveWorkbook
MainWB.Sheets.Add.Name = "PathSet"
Set Rng = Range("A1")

Application.DisplayAlerts = False

'define folder path
FolderPath = "C:\Users\santa\Desktop\fhb\revised_frontier_model\stock_dfs"

'count number of CSVs in folder

PathCountCondition = FolderPath & "\*.csv"

FileName = Dir(PathCountCondition)

Do While FileName <> ""
    Rng.Value = FileName
    Rng.Offset(1, 0).Select
    count = count + 1
    Set Rng = ActiveCell
    FileName = Dir()
Loop

Worksheets("PathSet").Activate
Set Rng = Range("A1")
Rng.Select
Range(Selection, Selection.End(xlDown)).Select
count = Selection.Rows.count

Worksheets("PathSet").Activate
Rng.Select

For FileNumber = 1 To count 'you can change count to a constant for sample runs
    
    'open the file
    
    FileName = FolderPath & "\" & Rng
    
    Set WB = Workbooks.Open(FileName)
    
    'copy its contents
    
    WB.Activate
    Range("A1").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy

    'create new sheet, and paste it into the main workbook
    
    MainWB.Activate
    RngNoPath = Left(Rng, Len(Rng) - 4)
    MainWB.Sheets.Add.Name = RngNoPath & "(D)"
    Range("A1").Select
    ActiveSheet.Paste
    Selection.Columns.AutoFit
    Range("A1").Select
    
    'close file
    WB.Close
    
    Call orderDataForGraphingIEX
    Call manipulateDataIEX
    'Call monthlyData(RngNoPath)
    
    'Worksheets("PathSet").Activate
    
    Worksheets("PathSet").Activate
    Rng.Offset(1, 0).Select
    Set Rng = ActiveCell
    
Next FileNumber

Worksheets("PathSet").Delete
MainWB.Sheets.Add.Name = "MonthSummary"
MainWB.Sheets.Add.Name = "DailySummary"

Worksheets("MonthSummary").Activate
Range("A1").Value = "Ticker"
Range("A2").Value = "Average"
Range("A3").Value = "Variance"
Range("A4").Value = "StrdDev"

For Each CurrentSheet In Worksheets
    If InStr(1, CurrentSheet.Name, "(Mon)") > 0 Then
        CurrentSheet.Activate
        SheetName = Split(CurrentSheet.Name, "(")(0)
        Call monthlySummary(SheetName)
        CurrentSheet.Activate
    End If
    'If InStr(1, CurrentSheet.Name, "(D)") > 0 Then
    '    CurrentSheet.Activate
    '    SheetName = CurrentSheet.Name
    '    Call dailySummary(SheetName)
    '    CurrentSheet.Activate
    'End If
Next
                                        
'tell me how long it took
SecondsElapsed = Round(Timer - StartTime, 2)
tickersPerSec = Round(SecondsElapsed / count, 2)
MsgBox "This code ran successfully in " & SecondsElapsed & " seconds" & vbCrLf & "Approximately " & tickersPerSec & " seconds per ticker", vbInformation
                                        
End Sub
Function orderDataForGraphingIEX()
    
    'order the columns for graphing
                                                            
    Columns("B:B").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Columns("G:G").Select
    Selection.Cut
    Columns("B:B").Select
    ActiveSheet.Paste
    
    'add commas and dollar signs
    
    Range("B2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Style = "Comma"
    Selection.NumberFormat = "_(* #,##0_);_(* (#,##0);_(* ""-""??_);_(@_)"
    Range("C2").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Style = "Currency"
    
End Function
Function manipulateDataIEX()
    Dim Rng As Range
    Dim LastRow As Integer
    
    'find the last row
    
    Range("A1").Select
    Range(Selection, Selection.End(xlDown)).Select
    LastRow = Selection.Rows.count
    
    'day average average
    Range("G1").Value = "Day Average"
    Range("G2").Select
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = "=AVERAGE(RC[-4]:RC[-1])"
    Range("G2").Select
    Selection.AutoFill Destination:=Range("G2:G" & LastRow)
    
    'data manipulations
    Range("H1").Value = "Intraday Open to Close"
    Application.CutCopyMode = False
    Range("H2").Select
    ActiveCell.FormulaR1C1 = "=RC[-2]-RC[-5]"
    Range("I1").Value = "Intraday %"
    Range("I2").Select
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = "=RC[-1]/RC[-6]"
    
    Range("H2:I2").Select
    Selection.AutoFill Destination:=Range("H2:I" & LastRow)
    Range("H:H").Select
    Selection.Style = "Currency"
    Range("I:I").Select
    Selection.Style = "Percent"
    Selection.NumberFormat = "0.000%"
    
    
    'resize one last time
    Range("A:I").Select
    Selection.Columns.AutoFit
    
    'set active cell to home
    Range("A1").Select
   
End Function
