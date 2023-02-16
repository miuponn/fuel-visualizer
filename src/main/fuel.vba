Sub CreateFuelSheet()

'Declare variables
Dim ws As Worksheet, ps As Worksheet, data As Worksheet
Dim DateCell As String, DateYear As String, DateMonth As String

Set ps = Sheets("Monthly Output By Fuel")
Set data = Sheets("Data")
DateCell = data.Range("A3")
DateYear = Right(DateCell, 4)

On Error Resume Next
Set ws = Worksheets(DateYear & " Output By Fuel Type")
If Err.Number = 9 Then
    Set ws = Worksheets.Add(Before:=Sheets("Monthly Output"))
    ws.name = DateYear & " Output By Fuel Type"
End If

ws.Tab.ColorIndex = 34

With ws
.Columns("A:A").ColumnWidth = 32
.Columns("A:N").ColumnWidth = 14
.Range("A1").Value = "Fuel Type"
.Range("B1").Value = "Jan" & DateYear
.Range("C1").Value = "Feb" & DateYear
.Range("D1").Value = "Mar" & DateYear
.Range("E1").Value = "Apr" & DateYear
.Range("F1").Value = "May" & DateYear
.Range("G1").Value = "Jun" & DateYear
.Range("H1").Value = "Jul" & DateYear
.Range("I1").Value = "Aug" & DateYear
.Range("J1").Value = "Sep" & DateYear
.Range("K1").Value = "Oct" & DateYear
.Range("L1").Value = "Nov" & DateYear
.Range("M1").Value = "Dec" & DateYear
.Range("N1").Value = "Annual Sum"
Range("A1:N1").Font.Bold = True
End With

'Find column of month
data.Cells(5, 1).Copy Destination:=ws.Cells(1, 15)
ws.Range("B1:M1").NumberFormat = "General"
ws.Cells(1, 15).NumberFormat = "General"
DateMonth = ws.Cells(1, 15).Value

Dim rng1 As Range
With ws.Rows(1)
    Set rng1 = .Find(What:=DateMonth)
End With

ws.Range("B1:M1").NumberFormat = "mmm-yy;@"
ws.Cells(1, 15).Clear

'Copy and paste fuel list
Dim rng2 As Range, rng3 As Range

Set rng2 = ps.PivotTables(1).PivotFields("Fuel Type").DataRange
Set rng3 = ws.Range("A" & ws.Rows.Count).End(xlUp).Offset(1, 0) 'Paste range starting from A2 and then first empty cell

rng2.Copy Destination:=rng3 'Copy/Paste
ws.Columns(1).Select
Selection.Interior.ColorIndex = xlNone

'Remove duplicates & "output" row
Dim rng4 As Range

Set rng4 = ws.Range("A2:A" & Cells(Rows.Count, 1).End(xlUp).Row)
ws.Columns(1).RemoveDuplicates (1)

rng4.Font.Bold = False

For Each cell In rng4
    If cell.Value = "Output" Then
        cell.EntireRow.Delete
    End If
Next

'Fill table with 0 MW
Dim tbl As Range

Set tbl = ws.Range("B2:N" & Cells(Rows.Count, 1).End(xlUp).Row)

For Each cell In tbl
    If IsEmpty(cell) = True Then
        cell.Value = "0"
    End If
Next
ws.Rows((ws.Range("a" & ws.Rows.Count).End(xlUp).Offset(1, 0).Row) & ":" & ws.Rows.Count).Delete

'Find and set fuels' MW
Dim fuel As String, mw As String, fuel2 As Range

For Each cell In rng4
    fuel = cell.Text
    mw = ps.Cells.Find(fuel).Offset(0, 1).Value
    cell.Offset(0, rng1.Column - 1).Select
    Selection.Value = mw
Next

'Calculate Annual Sum
Dim r As Range
Set r = ws.Range("B2:M" & Range("B" & Rows.Count).End(xlUp).Row)
Dim Total As Range: Set Total = r.Offset(, 12).Resize(r.Rows.Count, 1)

With Total
    .FormulaR1C1 = "=SUM(RC[-12]:RC[-1])"
    .Value = .Value
End With

ws.Range("N2:N" & Cells(Rows.Count, 1).End(xlUp).Row).NumberFormat = "General"

'Format numbers
ws.Range("C2:O" & Range("C" & Rows.Count).End(xlUp).Row).NumberFormat = "#,##0"

'Zoom window to 80%
ActiveWindow.Zoom = 80

'Format numbers
ws.Range("B2:N" & Range("C" & Rows.Count).End(xlUp).Row).NumberFormat = "#,##0"

'Bold annual sum column
ws.Columns(14).Font.Bold = True

'Shade every other gen
Dim Counter As Integer
    ws.Range("A2:N" & Cells(Rows.Count, 1).End(xlUp).Row).Select
   'For every row in the current selection...
    For Counter = 1 To Selection.Rows.Count
        'If the row is an odd number (within the selection)...
        If Counter Mod 2 = 1 Then
            'Shade rows
            Selection.Rows(Counter).Interior.Color = RGB(222, 235, 247)
        End If
    Next
    
'Get rid of previous borders
ws.Cells.Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    Selection.Borders(xlEdgeLeft).LineStyle = xlNone
    Selection.Borders(xlEdgeTop).LineStyle = xlNone
    Selection.Borders(xlEdgeBottom).LineStyle = xlNone
    Selection.Borders(xlEdgeRight).LineStyle = xlNone
    Selection.Borders(xlInsideVertical).LineStyle = xlNone
    Selection.Borders(xlInsideHorizontal).LineStyle = xlNone

'Add borders to every gen
ws.Range("A2:N" & Cells(Rows.Count, 1).End(xlUp).Row).Select
    'For every row in the current selection...
For Counter = 1 To Selection.Rows.Count
 'Shade rows
    With Selection
        .Rows(Counter).Borders.LineStyle = xlContinuous
        .Rows(Counter).Borders(xlInsideVertical).LineStyle = xlNone
    End With
Next
        
'Clean excess
For Each cell In ws.Range("A" & Cells(Rows.Count, 1).End(xlUp).Row.Offset(1, 0) & ":N" & Cells(Rows.Count, 1).End(xlUp).Row.Offset(1, 0))
    If cell.IsEmpty = False Then
        cell.EntireRow.Delete
    End If
Next

'Freeze top row
ActiveWindow.FreezePanes = False
Application.ScreenUpdating = False

Rows("2:2").Select
ActiveWindow.FreezePanes = True

ActiveWindow.FreezePanes = True

ws.Columns(16).EntireColumn.Delete
ws.Columns(15).EntireColumn.Delete

'Delete chart if existing
On Error Resume Next
Application.DisplayAlerts = False
    Sheets(DateYear & " MWh By Fuel Type Chart").Delete
    Sheets(DateYear & " TWh By Fuel Type Chart").Delete
    Sheets(DateYear & " Output By Fuel Type (TWh)").Delete
Application.DisplayAlerts = True

'Create Stacked Chart
Dim MyChart As Chart
Set MyChart = Charts.Add
ActiveChart.name = DateYear & " MWh By Fuel Type Chart"

With MyChart
    .SetSourceData ws.UsedRange
    .ChartType = xlColumnStacked
    .HasTitle = True
    .ChartTitle.Text = DateYear & " Monthly Output By Fuel Type"
    .HasLegend = True
    .SeriesCollection(1).Interior.Color = RGB(112, 173, 71) 'green biofuel
    .SeriesCollection(2).Interior.Color = RGB(165, 165, 165) 'grey gas
    .SeriesCollection(3).Interior.Color = RGB(91, 155, 213) 'blue hydro
    .SeriesCollection(4).Interior.Color = RGB(0, 43, 130) 'indigo nuclear
    .SeriesCollection(5).Interior.Color = RGB(255, 192, 0) 'yellow solar
    .SeriesCollection(6).Interior.Color = RGB(222, 235, 247) 'light blue wind
    .Axes(xlValue, xlPrimary).HasTitle = True
    .Axes(xlValue, xlPrimary).AxisTitle.Characters.Text = "Monthly Output (MWh)"
End With

Sheets(DateYear & " MWh By Fuel Type Chart").Tab.ColorIndex = 24
Sheets(DateYear & " MWh By Fuel Type Chart").Move After:=ws

'Make worksheet in kWh
ws.Copy After:=ws
ActiveSheet.name = DateYear & " Output By Fuel Type (TWh)"
Set kw = Sheets(DateYear & " Output By Fuel Type (TWh)")

Dim rng6 As Range
Set rng6 = kw.Range("B2:N" & Cells(Rows.Count, 1).End(xlUp).Row)

For Each cell In rng6
    cell.Value = cell.Value / 1000000
    cell.NumberFormat = "#.00"
Next

'Create stacked chart in kWh
Dim MyChart2 As Chart
Set MyChart2 = Charts.Add
ActiveChart.name = DateYear & " TWh By Fuel Type Chart"

With MyChart2
    .SetSourceData kw.UsedRange
    .ChartType = xlColumnStacked
    .HasTitle = True
    .ChartTitle.Text = DateYear & " Monthly Output By Fuel Type"
    .HasLegend = True
    .SeriesCollection(1).Interior.Color = RGB(112, 173, 71) 'green biofuel
    .SeriesCollection(2).Interior.Color = RGB(165, 165, 165) 'grey gas
    .SeriesCollection(3).Interior.Color = RGB(91, 155, 213) 'blue hydro
    .SeriesCollection(4).Interior.Color = RGB(0, 43, 130) 'indigo nuclear
    .SeriesCollection(5).Interior.Color = RGB(255, 192, 0) 'yellow solar
    .SeriesCollection(6).Interior.Color = RGB(222, 235, 247) 'light blue wind
    .Axes(xlValue, xlPrimary).HasTitle = True
    .Axes(xlValue, xlPrimary).AxisTitle.Characters.Text = "Monthly Output (TWh)"
    .Axes(xlValue, xlPrimary).TickLabels.NumberFormat = "0"
End With

Sheets(DateYear & " TWh By Fuel Type Chart").Tab.ColorIndex = 24
Sheets(DateYear & " TWh By Fuel Type Chart").Move After:=Sheets(DateYear & " MWh By Fuel Type Chart")

ws.Cells(1, 16).Value = "Output values displayed in MWh"
kw.Cells(1, 16).Value = "Output values displayed in TWh"

End Sub
