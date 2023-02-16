Sub CreateTableSheet3()

Sheets("Data").Copy Before:=Sheets(Sheets.Count)
ActiveSheet.name = "Table"
Worksheets("Table").Rows("1:3").Delete

End Sub
Sub InsertGas()

'Declare Variables
Dim PSheet As Worksheet
Dim DSheet As Worksheet
Dim PCache As PivotCache
Dim PTable As PivotTable
Dim PRange As Range
Dim lastRow As Long
Dim LastCol As Long

'Add New Worksheet
On Error Resume Next
Application.DisplayAlerts = False
Worksheets("Gas").Delete
Sheets.Add Before:=ActiveSheet
ActiveSheet.name = "Gas"
Application.DisplayAlerts = True
Set PSheet = Worksheets("Gas")
Set DSheet = Worksheets("Table")

'Define Data Range
lastRow = DSheet.Cells(Rows.Count, 1).End(xlUp).Row
LastCol = DSheet.Cells(1, Columns.Count).End(xlToLeft).Column
Set PRange = DSheet.Cells(1, 1).Resize(lastRow, LastCol)

'Define Pivot Cache
Set PCache = ActiveWorkbook.PivotCaches.Create _
(SourceType:=xlDatabase, SourceData:=PRange). _
CreatePivotTable(TableDestination:=PSheet.Cells(1, 1), _
TableName:="GasTable")

'Insert Blank Pivot Table
Set PTable = PCache.CreatePivotTable _
(TableDestination:=PSheet.Cells(1, 1), TableName:="GasTable")

'Insert Row Fields
With ActiveSheet.PivotTables("GasTable").PivotFields("Fuel Type")
.Orientation = xlRowField
.Position = 1
End With

With ActiveSheet.PivotTables("GasTable").PivotFields("Generator")
.Orientation = xlRowField
.Position = 2
End With

'Insert Column Fields
With ActiveSheet.PivotTables("GasTable").PivotFields("Measurement")
.Orientation = xlColumnField
.Position = 1
End With

'Insert Calculated Field
With ActiveSheet.PivotTables("GasTable")
    .CalculatedFields.Add "MW", "='Hour 1'+'Hour 2'+'Hour 3'+'Hour 4'+'Hour 5'+'Hour 6'+'Hour 7'+'Hour 8'+'Hour 9'+'Hour 10'+'Hour 11'+'Hour 12'+'Hour 13'+'Hour 14'+'Hour 15'+'Hour 16'+'Hour 17'+'Hour 18'+'Hour 19'+'Hour 20'+'Hour 21'+'Hour 22'+'Hour 23'+'Hour 24'"
    .PivotFields("MW").Orientation = xlDataField
    .Position = 1
    .NumberFormat = "#,##0"""
End With

ActiveSheet.PivotTables("GasTable").RowGrand = False
ActiveSheet.PivotTables("GasTable").ColumnGrand = False

'Format Pivot Table
ActiveSheet.PivotTables("GasTable").ShowTableStyleRowStripes = True
ActiveSheet.PivotTables("GasTable").TableStyle2 = "PivotStyleLight6"

End Sub

Sub AddGasSlicer()

'Declare Variables
Dim ws As Worksheet
Dim wb As Workbook
Dim pt As PivotTable
Dim SLCache As SlicerCache
Dim SL As Slicer

Set wb = ActiveWorkbook
Set ws = Worksheets("Gas")
Set pt = ws.PivotTables("GasTable")

'Create Slicer Cache
Set SLCache = wb.SlicerCaches.Add2(pt, "Fuel Type", "FuelTypeSlicerCache", XlSlicerCacheType.xlSlicer)

'Create Slicer
Set SL = SLCache.Slicers.Add(ws, , "FuelTypeSlicer", "Select a Fuel Type")

'Define which values are not selected in the Slicer
SLCache.SlicerItems("SOLAR").Selected = False
SLCache.SlicerItems("HYDRO").Selected = False
SLCache.SlicerItems("BIOFUEL").Selected = False
SLCache.SlicerItems("WIND").Selected = False
SLCache.SlicerItems("NUCLEAR").Selected = False

'Create Slicer Cache 2
Dim SLCache2 As SlicerCache, SL2 As Slicer
Set SLCache2 = wb.SlicerCaches.Add2(pt, "Measurement", "MeasurementSlicerCache4", XlSlicerCacheType.xlSlicer)

'Create Slicer
Set SL2 = SLCache2.Slicers.Add(ws, , "MeasurementSlicer4", "Select a Measurement")

'Define which values are not selected in the Slicer
On Error Resume Next
SLCache2.SlicerItems("Forecast").Selected = False

End Sub

Sub DeleteTableSheet3()

Application.DisplayAlerts = False 'switching off the alert button
Sheets("Table").Delete
Application.DisplayAlerts = True 'switching on the alert button

End Sub

Sub RunGasReport()

'Run all macros with a call statement
Call CreateTableSheet3
Call InsertGas
Call AddGasSlicer
Call DeleteTableSheet3

End Sub

