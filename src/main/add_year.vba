Sub CreateSheet()

'Declare variables

Dim ws As Worksheet, ps As Worksheet, data As Worksheet
Dim DateCell As String, DateYear As String, DateMonth As String
Dim wb As Workbook, pb As Workbook
Dim name As String

Set ps = Sheets("Monthly Output")
Set data = Sheets("Data")
DateCell = data.Range("A3")
DateYear = Right(DateCell, 4)


On Error Resume Next
Set ws = Worksheets(DateYear & " Output By Generator")
If Err.Number = 9 Then
    Set ws = Worksheets.Add(Before:=Sheets("Monthly Output"))
    ws.name = DateYear & " Output By Generator"
End If

ws.Tab.ColorIndex = 34

With ws
.Columns("A:A").ColumnWidth = 32
.Columns("B:B").ColumnWidth = 16
.Columns("O:O").ColumnWidth = 11
.Range("A1").Value = "Generator"
.Range("B1").Value = "Fuel Type"
.Range("C1").Value = "Jan" & DateYear
.Range("D1").Value = "Feb" & DateYear
.Range("E1").Value = "Mar" & DateYear
.Range("F1").Value = "Apr" & DateYear
.Range("G1").Value = "May" & DateYear
.Range("H1").Value = "Jun" & DateYear
.Range("I1").Value = "Jul" & DateYear
.Range("J1").Value = "Aug" & DateYear
.Range("K1").Value = "Sep" & DateYear
.Range("L1").Value = "Oct" & DateYear
.Range("M1").Value = "Nov" & DateYear
.Range("N1").Value = "Dec" & DateYear
.Range("O1").Value = "Annual Sum"
Range("A1:O1").Font.Bold = True
End With

'Find column of month
data.Cells(5, 1).Copy Destination:=ws.Cells(1, 16)
ws.Range("C1:N1").NumberFormat = "General"
ws.Cells(1, 16).NumberFormat = "General"
DateMonth = ws.Cells(1, 16).Value

Dim rng1 As Range
With ws.Rows(1)
    Set rng1 = .Find(What:=DateMonth)
End With

ws.Range("C1:N1").NumberFormat = "mmm-yy;@"
ws.Cells(1, 16).Clear

'Copy and paste generator list
Dim rng2 As Range, rng3 As Range

Set rng2 = ps.PivotTables(1).PivotFields("Generator").DataRange
Set rng3 = ws.Range("A" & ws.Rows.Count).End(xlUp).Offset(1, 0) 'Paste range starting from A2 and then first empty cell

rng2.Copy Destination:=rng3 'Copy/Paste
ws.Columns(1).Select
Selection.Interior.ColorIndex = xlNone

'Remove duplicates
ws.Columns(1).RemoveDuplicates (1)

'Fill table with 0 MW
Dim tbl As Range

Set tbl = ws.Range("C2:N" & Cells(Rows.Count, 1).End(xlUp).Row)

For Each Cell In tbl
    If IsEmpty(Cell) = True Then
        Cell.Value = "0"
    End If
Next
ws.Rows((ws.Range("a" & ws.Rows.Count).End(xlUp).Offset(1, 0).Row) & ":" & ws.Rows.Count).Delete
    
'Find gens' MW
Dim gen As String, fuel As String, mw As String, gen2 As Range


'Set gens' MW & Fuel Type
For Each Cell In rng2
    gen = Cell.Text
    mw = Cell.Offset(0, 1).Text
    ws.Cells.Find(gen).Offset(0, rng1.Column - 1).Select
    Selection.Value = mw
    fuel = data.Cells.Find(gen).Offset(0, 1).Text
    ws.Cells.Find(gen).Offset(0, 1).Value = fuel
Next

'Calculate Annual Sum
Dim r As Range
Set r = ws.Range("C2:N" & Range("C" & Rows.Count).End(xlUp).Row)
Dim Total As Range: Set Total = r.Offset(, 12).Resize(r.Rows.Count, 1)

With Total
    .FormulaR1C1 = "=SUM(RC[-12]:RC[-1])"
    .Value = .Value
End With

'Format numbers
ws.Range("C2:O" & Range("C" & Rows.Count).End(xlUp).Row).NumberFormat = "#,##0"

'Zoom window to 80%
ActiveWindow.Zoom = 80

'Bold annual sum column
ws.Columns(15).Font.Bold = True

'Make cells color invisible
ws.Cells.Select
Selection.Interior.ColorIndex = xlNone

'Shade every other gen
Dim Counter As Integer
    ws.Range("A2:O" & Cells(Rows.Count, 1).End(xlUp).Row).Select
   'For every row in the current selection...
    For Counter = 1 To Selection.Rows.Count
        'If the row is an odd number (within the selection)...
        If Counter Mod 2 = 1 Then
            'Shade rows
            Selection.Rows(Counter).Interior.Color = RGB(222, 235, 247)
        End If
    Next
    
'Freeze top row
With ActiveWindow
    If .FreezePanes Then .FreezePanes = False
    .ScrollRow = 1
    .ScrollColumn = 1
    .SplitColumn = 0
    .SplitRow = 1
    .FreezePanes = True
End With

ws.Cells(1, 15).Value = "Annual Sum"

'Add units of measurement
ws.Cells(1, 16).Value = "Output values displayed in MWh"

End Sub

