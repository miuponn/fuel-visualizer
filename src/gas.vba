Sub CreateGasSheet()

'Declare variables
Dim ws As Worksheet, ps As Worksheet, data As Worksheet
Dim DateCell As String, DateYear As String, DateMonth As String

Set ps = Sheets("Gas")
Set data = Sheets("Data")
DateCell = data.Range("A3")
DateYear = Right(DateCell, 4)

On Error Resume Next
Set ws = Worksheets(DateYear & " Gas Measurements")
If Err.Number = 9 Then
    Set ws = Worksheets.Add(Before:=Sheets("Monthly Output"))
    ws.name = DateYear & " Gas Measurements"
End If

ws.Tab.ColorIndex = 34

On Error Resume Next
ws.Columns(1).UnMerge

With ws
.Columns("A:A").ColumnWidth = 32
.Columns("B:B").ColumnWidth = 16
.Range("A1").Value = "Generator"
.Range("B1").Value = "Measurement"
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
.Range("A1:O1").Font.Bold = True
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

'Cut and paste existing MW
Dim mwcut As Worksheet
Dim rng5 As Range

Set mwcut = Sheets.Add
    mwcut.name = "mwcut"
Set rng5 = ws.UsedRange

rng5.Copy Destination:=mwcut.Range("A" & ws.Rows.Count).End(xlUp) 'Cut/Paste

'Remove duplicates & extra row
Dim rng4 As Range

Set rng4 = ws.Range("A2:A" & Cells(Rows.Count, 1).End(xlUp).Row)
ws.Columns(1).RemoveDuplicates (1)

rng4.Font.Bold = False

For Each Cell In rng4
    If Cell.Value = "GAS" Then
        Cell.EntireRow.Delete
    End If
Next

'Add blank rows
Dim lastRow As Long, RowNumber As Long
With ws
lastRow = .Cells(Rows.Count, "A").End(xlUp).Row
    For RowNumber = lastRow To 3 Step -1
        .Rows(RowNumber).Insert
        .Rows(RowNumber).Insert
    Next RowNumber
End With

If IsEmpty(ws.Cells(2, 1).Offset(3, 0)) = True Then
    ws.Cells(2, 1).Offset(5, 0).EntireRow.Delete
    ws.Cells(2, 1).Offset(4, 0).EntireRow.Delete
    ws.Cells(2, 1).Offset(3, 0).EntireRow.Delete
End If

ws.Range("C2:N" & Cells(Rows.Count, 1).End(xlUp).Row).ClearContents

'Cut and paste existing MW
mwcut.Range("C:N").Copy ws.Range("C:N")

Application.DisplayAlerts = False

Sheets("mwcut").Delete

Application.DisplayAlerts = True

'Add measurements
Dim capmes As String, outmes As String, cfmes As String

For Each Cell In rng4
    If IsEmpty(Cell) = False Then
        capmes = "Capability (MW)"
        outmes = "Output (MWh)"
        cfmes = "CF (%)"
        Cell.Offset(0, 1).Select
        Selection.Value = capmes
        Cell.Offset(1, 1).Select
        Selection.Value = outmes
        Cell.Offset(2, 1).Select
        Selection.Value = cfmes
    End If
Next
ws.Rows((ws.Range("a" & ws.Rows.Count).End(xlUp).Offset(3, 0).Row) & ":" & ws.Rows.Count).Delete

'Fill table with 0 MW
Dim tbl As Range

Set tbl = ws.Range("C2:N" & ws.Range("b" & ws.Rows.Count).End(xlUp).Row)
For Each Cell In tbl
    If IsEmpty(Cell) = True Then
        Cell.Value = "0"
    End If
Next

ws.Rows((ws.Range("b" & ws.Rows.Count).End(xlUp).Offset(1, 0).Row) & ":" & ws.Rows.Count).Delete

'Find gens' MW
Dim gen As String, cap As String, out As String, cf As String, gen2 As Range

For Each Cell In rng4
    If IsEmpty(Cell) = False Then
        gen = Cell.Text
        cap = ps.Cells.Find(gen).Offset(0, 1).Value
        out = ps.Cells.Find(gen).Offset(0, 2).Value
        Cell.Offset(0, rng1.Column - 1).Select
        Selection.Value = cap
        Selection.NumberFormat = "#,##0"
        Cell.Offset(1, rng1.Column - 1).Select
        Selection.Value = out
        Selection.NumberFormat = "#,##0"
        Cell.Offset(2, rng1.Column - 1).Select
        Selection.Value = out / cap
        Selection.NumberFormat = "0.0%"
    End If
Next

'Calculate Annual Sum
Dim r As Range
Set r = ws.Range("C2:N" & Range("C" & Rows.Count).End(xlUp).Row)
Dim Total As Range: Set Total = r.Offset(, 12).Resize(r.Rows.Count, 1)

With Total
    .FormulaR1C1 = "=SUM(RC[-12]:RC[-1])"
    .Value = .Value
End With

ws.Range("O2:O" & Cells(Rows.Count, 1).End(xlUp).Row).NumberFormat = "#,##0"

'Calculate annual CF
For i = 4 To Range("N" & Rows.Count).End(xlUp).Row Step 3
    ws.Cells(i, 15).Value = ws.Cells(i, 15).Offset(-1, 0).Value / ws.Cells(i, 15).Offset(-2, 0).Value
    ws.Cells(i, 15).NumberFormat = "0.0%"
Next i

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
        If Counter Mod 6 = 1 Then
            'Shade rows
            Selection.Rows(Counter).Interior.Color = RGB(222, 235, 247)
            Selection.Rows(Counter).Offset(1, 0).Interior.Color = RGB(222, 235, 247)
            Selection.Rows(Counter).Offset(2, 0).Interior.Color = RGB(222, 235, 247)
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

'Add borders to every other gen
ws.Range("A2:O" & Cells(Rows.Count, 1).End(xlUp).Row).Select
'For every row in the current selection...
For Counter = 1 To Selection.Rows.Count
    'If the row is an odd number (within the selection)...
    If Counter Mod 3 = 1 Then
        With Selection
            .Rows(Counter).Borders(xlDiagonalDown).LineStyle = xlNone
            .Rows(Counter).Borders(xlDiagonalUp).LineStyle = xlNone
        End With
        With Selection.Rows(Counter).Borders(xlEdgeLeft)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlThin
        End With
            With Selection.Rows(Counter).Borders(xlEdgeRight)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlThin
        End With
        With Selection.Rows(Counter).Borders(xlEdgeTop)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlThin
        End With
        
    End If
Next

ws.Range("A3:O" & Cells(Rows.Count, 1).End(xlUp).Row).Select
For Counter = 1 To Selection.Rows.Count
    If Counter Mod 3 = 1 Then
        With Selection
            .Rows(Counter).Borders(xlDiagonalDown).LineStyle = xlNone
            .Rows(Counter).Borders(xlDiagonalUp).LineStyle = xlNone
        End With
        With Selection.Rows(Counter).Borders(xlEdgeLeft)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlThin
        End With
            With Selection.Rows(Counter).Borders(xlEdgeRight)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlThin
        End With
    End If
Next
   
ws.Range("A4:O" & Cells(Rows.Count, 1).End(xlUp).Offset(2, 0).Row).Select
For Counter = 1 To Selection.Rows.Count
    If Counter Mod 3 = 1 Then
        With Selection
            .Rows(Counter).Borders(xlDiagonalDown).LineStyle = xlNone
            .Rows(Counter).Borders(xlDiagonalUp).LineStyle = xlNone
        End With
        With Selection.Rows(Counter).Borders(xlEdgeLeft)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlThin
        End With
            With Selection.Rows(Counter).Borders(xlEdgeRight)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlThin
        End With
        With Selection.Rows(Counter).Borders(xlEdgeBottom)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlThin
        End With
    End If
Next

'Merge gens
Application.DisplayAlerts = False

For i = 2 To Range("A" & Rows.Count).End(xlUp).Row Step 3
    With Range("A" & i).Resize(3, 1)
        .MergeCells = True
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
Next i

Application.DisplayAlerts = True

'Freeze top row
ActiveWindow.FreezePanes = False
Application.ScreenUpdating = False

Rows("2:2").Select
ActiveWindow.FreezePanes = True

ActiveWindow.FreezePanes = True

End Sub

