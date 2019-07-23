Attribute VB_Name = "Module1"
Sub Button1_Click()
Dim PSheet As Worksheet
Dim DSheet As Worksheet
Dim PCache As PivotCache
Dim PTable As pivotTable
Dim PRange As Range
Dim LastRow As Long
Dim LastCol As Long
Dim rawData As String

rawData = "riskReport07222019.csv"

Workbooks.Open (ThisWorkbook.Path & "\" & rawData)

'Create new workbook to hold report
Set New_wkb = Workbooks.Add
timeNow = Now()
RightNow = Format(timeNow, "mmmm dd, yyyy")
thisReport = "Sherpa Report" & "-" & RightNow & ".xlsm"
sherpaReportGen = "SherpaReportGenerator.xlsm"

With New_wkb
    .Title = "Sherpa Report" & "-" & DateTime.Now
    .Subject = "Custom Sherpa Report built" & "-" & DateTime.Now
    .SaveAs Filename:=ThisWorkbook.Path & "\" & thisReport, FileFormat:=xlOpenXMLWorkbookMacroEnabled
End With

'Create Headings
Workbooks(thisReport).Sheets(1).Cells(1, 1) = "Sales Rep"
Workbooks(thisReport).Sheets(1).Cells(1, 2) = "Team"
Workbooks(thisReport).Sheets(1).Cells(1, 3) = "Customer Name"
Workbooks(thisReport).Sheets(1).Cells(1, 4) = "Lease Number"
Workbooks(thisReport).Sheets(1).Cells(1, 5) = "Maturity Date"
Workbooks(thisReport).Sheets(1).Cells(1, 6) = "Equipment Payment"
Workbooks(thisReport).Sheets(1).Cells(1, 7) = "Lease Provider"
Workbooks(thisReport).Sheets(1).Cells(1, 8) = "Total Funded"
Workbooks(thisReport).Sheets(1).Cells(1, 9) = "Volume Level"
Workbooks(thisReport).Sheets(1).Cells(1, 10) = "Model"
Workbooks(thisReport).Sheets(1).Cells(1, 11) = "Serial Number"
Workbooks(thisReport).Sheets(1).Cells(1, 12) = "Address"

'count data cells
LastRow = Workbooks(rawData).Sheets(1).Cells(Rows.Count, "F").End(xlUp).Row

For i = 1 To LastRow
    'Fill in Sales Rep
    Workbooks(thisReport).Sheets(1).Cells(i + 1, 1).Value = Workbooks(rawData).Sheets(1).Cells(i, 6).Value
    'Customer Name
    Workbooks(thisReport).Sheets(1).Cells(i + 1, 3).Value = Workbooks(rawData).Sheets(1).Cells(i, 7).Value
    'Lease Number
    Workbooks(thisReport).Sheets(1).Cells(i + 1, 4).Value = Workbooks(rawData).Sheets(1).Cells(i, 9).Value
    'Maturity Date
    Workbooks(thisReport).Sheets(1).Cells(i + 1, 5).Value = Workbooks(rawData).Sheets(1).Cells(i, 11).Value
    'Equipment Payment
    Workbooks(thisReport).Sheets(1).Cells(i + 1, 6).Value = Workbooks(rawData).Sheets(1).Cells(i, 13).Value
    'Lease Provider
    Workbooks(thisReport).Sheets(1).Cells(i + 1, 7).Value = Workbooks(rawData).Sheets(1).Cells(i, 15).Value
    'Total Funded
    Workbooks(thisReport).Sheets(1).Cells(i + 1, 8).Value = Workbooks(rawData).Sheets(1).Cells(i, 17).Value
    'Model
    Workbooks(thisReport).Sheets(1).Cells(i + 1, 10).Value = Workbooks(rawData).Sheets(1).Cells(i, 27).Value
    'Serial Number
    Workbooks(thisReport).Sheets(1).Cells(i + 1, 11).Value = Workbooks(rawData).Sheets(1).Cells(i, 28).Value
    'Address
    Workbooks(thisReport).Sheets(1).Cells(i + 1, 12).Value = Workbooks(rawData).Sheets(1).Cells(i, 30).Value
Next i

totalRows = Workbooks(thisReport).Sheets(1).Cells(Rows.Count, "F").End(xlUp).Row
totalInLegend = Workbooks(sherpaReportGen).Sheets(2).Cells(Rows.Count, "A").End(xlUp).Row
totalInTeams = Workbooks(sherpaReportGen).Sheets(3).Cells(Rows.Count, "A").End(xlUp).Row

For i = 2 To totalInLegend
    'Set Machine to check
    machineToCheck = Workbooks(sherpaReportGen).Sheets(2).Cells(i, 1).Value
    volumeLevel = Workbooks(sherpaReportGen).Sheets(2).Cells(i, 2).Value
    
    'Look through column J for a match
    For j = 2 To totalRows
        If Workbooks(thisReport).Sheets(1).Cells(j, 10).Value = machineToCheck Then
            Workbooks(thisReport).Sheets(1).Cells(j, 9).Value = volumeLevel
        End If
    Next j
Next i

For i = 2 To totalInTeams
    repNameInLegend = Workbooks(sherpaReportGen).Sheets(3).Cells(i, 1).Value
    repTeam = Workbooks(sherpaReportGen).Sheets(3).Cells(i, 2).Value

    For j = 2 To totalRows
        If Workbooks(thisReport).Sheets(1).Cells(j, 1).Value = repNameInLegend Then
            Workbooks(thisReport).Sheets(1).Cells(j, 2).Value = repTeam
        End If
    Next j

Next i


Workbooks(thisReport).Sheets(1).Columns("A:R").AutoFit
Workbooks(thisReport).Sheets(1).Name = "Raw Data"
Workbooks(thisReport).Sheets(1).Tab.Color = 1

'Create a sheet for the Pivot table

On Error Resume Next
Application.DisplayAlerts = False
Worksheets("PivotTable").Delete
Sheets.Add Before:=ActiveSheet
ActiveSheet.Name = "PivotTable"
Application.DisplayAlerts = True
Set PSheet = Worksheets("PivotTable")
Set DSheet = Worksheets("Raw Data")

'Define Data range
LastRow = DSheet.Cells(Rows.Count, 1).End(xlUp).Row
LastCol = DSheet.Cells(1, Columns.Count).End(xlToLeft).Column
Set PRange = DSheet.Cells(1, 1).Resize(LastRow, LastCol)

'Define Pivot Cache
Set PCache = ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=PRange).CreatePivotTable(TableDestination:=PSheet.Cells(2, 2), TableName:="SalesPivotTable")

'Insert Blank Pivot Table
Set PTable = PCache.CreatePivotTable(TableDestination:=PSheet.Cells(1, 1), TableName:="SalesPivotTable")

'Insert Row Fields
With ActiveSheet.PivotTables("SalesPivotTable").PivotFields("Volume Level")
.Orientation = xlRowField
.Position = 1
End With

With ActiveSheet.PivotTables("SalesPivotTable").PivotFields("Total Funded")
.Orientation = xlDataField
.Position = 2
.NumberFormat = "$ #,##0"
End With

ActiveSheet.PivotTables("SalesPivotTable").PivotFields("Volume Level").PivotItems("3").Position = 4
ActiveSheet.PivotTables("SalesPivotTable").PivotFields("Volume Level").PivotItems("2").Position = 3
ActiveSheet.PivotTables("SalesPivotTable").PivotFields("Volume Level").PivotItems("1").Position = 2
ActiveSheet.PivotTables("SalesPivotTable").PivotFields("Volume Level").PivotItems("4").Position = 1

'Rename Cells
Workbooks(thisReport).Sheets("PivotTable").Cells(2, 2).Value = "Volume Class"
Workbooks(thisReport).Sheets("PivotTable").Cells(6, 2).Value = "High"
Workbooks(thisReport).Sheets("PivotTable").Cells(5, 2).Value = "Medium"
Workbooks(thisReport).Sheets("PivotTable").Cells(4, 2).Value = "Low"
Workbooks(thisReport).Sheets("PivotTable").Cells(3, 2).Value = "Unknown"

Dim pivotTable As pivotTable
Set pivotTable = ActiveSheet.PivotTables(1)
pivotTable.PivotSelect ("My Pivot")
Charts.Add2
ActiveChart.Location Where:=xlLocationAsObject, Name:=pivotTable.Parent.Name
ActiveChart.ChartType = xlBarClustered
ActiveChart.ChartStyle = 222
ActiveChart.ChartTitle.Delete
ActiveChart.ShowAllFieldButtons = False

'deselect chart
Range("A1").Select
Application.ScreenUpdating = True



End Sub
