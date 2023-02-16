ActiveCell.FormulaR1C1 = "=SUM(RC[-12]:RC[-1])"

.FormulaR1C1 = "=SUM(RC[-12]:RC[-2])"
    .Value = .Value

Dim r As Range:     Set r = ws.Range("C2:N" & Range("C" & Rows.Count).End(xlUp).Row)
Dim Total As Range: Set Total = r.Offset(, 12).Resize(r.Rows.Count, 1)

With Total
    .FormulaR1C1 = "=SUM(RC[-12]:RC[-2])"
    .Value = .Value
End With

Dim r As Range: Set r = ws.Range("B2:M" & Range("B" & Rows.Count).End(xlUp).Row)
Dim Total As Range: Set Total = r.Offset(, 12).Resize(r.Rows.Count, 1)

With Total
    .FormulaR1C1 = "=SUM(RC[-12]:RC[-2])"
    .Value = .Value
End With

ws.Range("N2:N" & Cells(Rows.Count, 1).End(xlUp).Row).NumberFormat = "General"

End Sub


Sub SumAlongOneRow()
'Declare a Long type variable for the next available row
'where all the SUM formulas will go.
'Declare a Long type variable to identify the last column
'in the used range.
Dim NextRow As Long, LastColumn As Long
NextRow = _
Cells.Find(What:="*", After:=Range("A1"), _
SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row + 1
LastColumn = _
Cells.Find(What:="*", After:=Range("A1"), _
SearchOrder:=xlByColumns, SearchDirection:=xlPrevious).Column
'The used range starts in column A which is Column 1 in VBA.
'The sales numbers in the table start on row 5.
'Therefore, sum the numbers in a single formula that starts from
'row 5 and ends at the next available row.
Range(Cells(NextRow, 1), Cells(NextRow, LastColumn)).FormulaR1C1 _
= "=SUM(R5C:R" & NextRow - 1 & "C)"
End Sub
