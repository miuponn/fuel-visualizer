Sub CreateTableSheet()

Sheets("Data").Copy Before:=Sheets(Sheets.Count)
ActiveSheet.name = "Table"
Worksheets("Table").Rows("1:3").Delete

End Sub
Sub InsertMonthlyOutput()

'Declare Variables
Dim PSheet As Worksheet, DSheet As Worksheet
Dim PCache As PivotCache, PTable As PivotTable, PRange As Range
Dim lastRow As Long, LastCol As Long

'Add New Worksheet
On Error Resume Next
Application.DisplayAlerts = False
Worksheets("Monthly Output").Delete
Sheets.Add Before:=ActiveSheet
ActiveSheet.name = "Monthly Output"
Application.DisplayAlerts = True
Set PSheet = Worksheets("Monthly Output")
Set DSheet = Worksheets("Table")

'Define Data Range
lastRow = DSheet.Cells(Rows.Count, 1).End(xlUp).Row
LastCol = DSheet.Cells(1, Columns.Count).End(xlToLeft).Column
Set PRange = DSheet.Cells(1, 1).Resize(lastRow, LastCol)

'Define Pivot Cache
Set PCache = ActiveWorkbook.PivotCaches.Create _
(SourceType:=xlDatabase, SourceData:=PRange). _
CreatePivotTable(TableDestination:=PSheet.Cells(1, 1), _
TableName:="MonthlyPivotTable")

'Insert Blank Pivot Table
Set PTable = PCache.CreatePivotTable _
(TableDestination:=PSheet.Cells(1, 1), TableName:="MonthlyPivotTable")

'Insert Row Fields
With ActiveSheet.PivotTables("MonthlyPivotTable").PivotFields("Generator")
.Orientation = xlRowField
.Position = 1
End With

'Insert Calculated Field
With ActiveSheet.PivotTables("MonthlyPivotTable")
    .CalculatedFields.Add "MW", "='Hour 1'+'Hour 2'+'Hour 3'+'Hour 4'+'Hour 5'+'Hour 6'+'Hour 7'+'Hour 8'+'Hour 9'+'Hour 10'+'Hour 11'+'Hour 12'+'Hour 13'+'Hour 14'+'Hour 15'+'Hour 16'+'Hour 17'+'Hour 18'+'Hour 19'+'Hour 20'+'Hour 21'+'Hour 22'+'Hour 23'+'Hour 24'"
    .PivotFields("MW").Orientation = xlDataField
    .Position = 1
    .NumberFormat = "#,##0"""
End With

ActiveSheet.PivotTables("MonthlyPivotTable").RowGrand = False
ActiveSheet.PivotTables("MonthlyPivotTable").ColumnGrand = False

'Format Pivot Table
ActiveSheet.PivotTables("MonthlyPivotTable").ShowTableStyleRowStripes = True
ActiveSheet.PivotTables("MonthlyPivotTable").TableStyle2 = "PivotStyleLight6"

End Sub
Sub AddSlicer()

'Declare Variables
Dim ws As Worksheet
Dim wb As Workbook
Dim pt As PivotTable
Dim SLCache As SlicerCache
Dim SL As Slicer

Set wb = ActiveWorkbook
Set ws = Worksheets("Monthly Output")
Set pt = ws.PivotTables("MonthlyPivotTable")

'Create Slicer Cache
Set SLCache = wb.SlicerCaches.Add2(pt, "Measurement", "MeasurementSlicerCache", XlSlicerCacheType.xlSlicer)

'Create Slicer
Set SL = SLCache.Slicers.Add(ws, , "MeasurementSlicer", "Select a Measurement")

'Define which values are not selected in the Slicer
On Error Resume Next
SLCache.SlicerItems("Available Capacity").Selected = False
SLCache.SlicerItems("Capability").Selected = False
SLCache.SlicerItems("Forecast").Selected = False


End Sub
Sub DeleteTableSheet()

Application.DisplayAlerts = False 'switching off the alert button
Sheets("Table").Delete
Application.DisplayAlerts = True 'switching on the alert button

End Sub

Sub RunMonthlyReport()

'Run all macros with a call statement
Call CreateTableSheet
Call InsertMonthlyOutput
Call AddSlicer
Call DeleteTableSheet

End Sub

