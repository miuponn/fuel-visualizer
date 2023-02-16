Dim rng5 As Range, rng6 As Range

Set rng5 = ws.Range("A2:A" & Cells(Rows.Count, 1).End(xlUp).Row)
For Each Cell In rng5
    If IsEmpty(Cell) = False Then
        Cell.Offset(0, 1).Value = "Capability"
    End If
Next

Set rng6 = ws.Range("B2:B" & Cells(Rows.Count, 1).End(xlUp).Row)
For Each Cell In rng6
    If IsEmpty(Cell.Offset(0, -1)) = True Then
        Cell.Value = "Pee"
    End If
Next

For Each Cell In rng6
    If Cell.Value = "Capability" Then
        If IsEmpty(Cell.Offset(0, -1)) = True Then
            Cell.Clear
        End If
    End If
Next

With ws
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
    Set rng1 = .Find(what:=DateMonth)
End With

ws.Range("C1:N1").NumberFormat = "mmm-yy;@"
ws.Cells(1, 16).Clear

'Copy and paste generator list
Dim rng2 As Range, rng3 As Range

Set rng2 = ps.PivotTables(1).PivotFields("Generator").DataRange
Set rng3 = ws.Range("A" & ws.Rows.Count).End(xlUp).Offset(1, 0) 'Paste range starting from A2 and then first empty cell

rng2.Copy Destination:=rng3 'Copy/Paste
ws.Columns(1).Select
Selection.Interior.Color = vbWhite

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

Dim fRg As Range
Dim lr, i As Long
lr = Cells(Rows.Count, "A").End(xlUp).Row 'Assumes rowdata existsin column A
Set fRg = Cells.Find(what:="NUCLEAR", lookat:=xlWhole)
For i = lr To (fRg.Row + 1) Step -1
    Cells(i, 1).EntireRow.Delete
Next i

For Each Cell In rng4
    If Cell.Value = "NUCLEAR" Then
        Cell.EntireRow.Delete
    End If
Next

'Add blank rows
Dim LastRow As Long, RowNumber As Long
With ws
LastRow = .Cells(Rows.Count, "A").End(xlUp).Row
    For RowNumber = LastRow To 3 Step -1
        .Rows(RowNumber).Insert
        .Rows(RowNumber).Insert
    Next RowNumber
End With

If IsEmpty(ws.Cells(2, 1).Offset(3, 0)) = True Then
    ws.Cells(2, 1).Offset(5, 0).EntireRow.Delete
    ws.Cells(2, 1).Offset(4, 0).EntireRow.Delete
    ws.Cells(2, 1).Offset(3, 0).EntireRow.Delete
End If

'Add measurements
Dim rng5 As Range

ws.Range("A:A").Copy
ws.Range("B:B").Insert

Set rng5 = ws.Range("B2:B" & Cells(Rows.Count, 1).End(xlUp).Row)
For Each Cell In rng5
    If IsEmpty(Cell) = False Then
        Cell.Value = "Capability"
    End If
Next

For Each Cell In rng5
    If Cell.Value = "Capability" Then
        Cell.Offset(1, 0).Value = "Output"
        Cell.Offset(2, 0).Value = "CF %"
    End If
Next

ws.Cells(1, 2).Value = "Measurement"
ws.Range("B1").Font.Bold = True

ws.Columns(3).Delete

'Fill table with 0 MW
Dim tbl As Range

Set tbl = ws.Range("C2:N" & Cells(Rows.Count, 1).End(xlUp).Row)

For Each Cell In tbl
    If IsEmpty(Cell) = True Then
        Cell.Value = "0"
    End If
Next

ws.Rows((ws.Range("b" & ActiveSheet.Rows.Count).End(xlUp).Offset(1, 0).Row) & ":" & ws.Rows.Count).Delete

'Find gens' MW
'Dim gen As String, cap As String, out As String, cf As String, gen2 As Range

'For Each Cell In rng4
    'If IsEmpty(Cell) = False Then
        'gen = Cell.Text
        'cap = ps.Cells.Find(gen).Offset(0, 1).Value
        'out = ps.Cells.Find(gen).Offset(0, 2).Value
        'Cell.Offset(0, rng1.Column - 1).Select
        'Selection.Value = cap
        'Cell.Offset(1, rng1.Column - 1).Select
        'Selection.Value = out
    'End If
'Next

End Sub
