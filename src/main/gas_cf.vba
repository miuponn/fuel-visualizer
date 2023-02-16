Application.ScreenUpdating = False
Columns("B:B").SpecialCells(xlCellTypeBlanks).EntireRow.Delete
Application.ScreenUpdating = True
       
With gas
    .Columns("A:A").Copy
    .Range("D1").Select
    .Paste
    .Columns("B:B").Font.Bold = False
    .Cells(1, 2) = "Output (MWh)"
    .Cells(1, 2).Font.Bold = True
    .Cells(1, 5) = "CF (%)"
    .Cells(1, 5).Font.Bold = True
    .Columns("E:E").ColumnWidth = 20
End With

Dim rng6 As Range, rng7 As Range, rng8 As Range, rng9 As Range
Set rng6 = gas.Range("A2:B" & Cells(Rows.Count, 1).End(xlUp).Row)
Set rng7 = gas.Range("D2:E" & Cells(Rows.Count, 1).End(xlUp).Row)
Set rng8 = gas.Range("B2:B" & Cells(Rows.Count, 1).End(xlUp).Row)
Set rng9 = gas.Range("E2:E" & Cells(Rows.Count, 1).End(xlUp).Row)

rng7.Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    Selection.Borders(xlEdgeTop).LineStyle = xlNone
    Selection.Borders(xlEdgeBottom).LineStyle = xlNone
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    Selection.Borders(xlInsideVertical).LineStyle = xlNone
    Selection.NumberFormat = "0.0%"
   
gas.Sort.SortFields.Add2 Key:=rng8, SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortNormal
With gas.Sort
    .SetRange rng6
    .Header = xlNo
    .MatchCase = False
    .Orientation = xlTopToBottom
    .SortMethod = xlPinYin
    .Apply
End With

gas.Sort.SortFields.Clear
gas.Sort.SortFields.Add2 Key:=rng9, SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortNormal
With gas.Sort
    .SetRange rng7
    .Header = xlNo
    .MatchCase = False
    .Orientation = xlTopToBottom
    .SortMethod = xlPinYin
    .Apply
End With

gas.Cells.Select
Selection.Interior.ColorIndex = xlNone

'Add color scaling
Dim cs As ColorScale, cs2 As ColorScale

Set cs = rng8.FormatConditions.AddColorScale(ColorScaleType:=3)
With cs
    'the first color is red
    With .ColorScaleCriteria(1)
        .FormatColor.Color = RGB(248, 105, 107)
        .Type = xlConditionValueLowestValue
    End With
    'the second color is yellow set at percentile value 50
    With .ColorScaleCriteria(2)
        .FormatColor.Color = RGB(255, 235, 132)
        .Type = xlConditionValuePercentile
        .Value = 50
    End With
    'the third color is green
    With .ColorScaleCriteria(3)
        .FormatColor.Color = RGB(99, 190, 123)
        .Type = xlConditionValueHighestValue
    End With
End With

Set cs2 = rng9.FormatConditions.AddColorScale(ColorScaleType:=3)
With cs2
    'the first color is red
    With .ColorScaleCriteria(1)
        .FormatColor.Color = RGB(248, 105, 107)
        .Type = xlConditionValueNumber
        .Value = 0
    End With
    'the second color is yellow set at percentile value 50
    With .ColorScaleCriteria(2)
        .FormatColor.Color = RGB(255, 235, 132)
        .Type = xlConditionValueNumber
        .Value = 0.5
    End With
    'the third color is green
    With .ColorScaleCriteria(3)
        .FormatColor.Color = RGB(99, 190, 123)
        .Type = xlConditionValueNumber
        .Value = 1
    End With
End With
    
'Merge gen cells
Application.DisplayAlerts = False

For i = 2 To ws.Range("A" & Rows.Count).End(xlUp).Row Step 3
    With ws.Range("A" & i).Resize(3, 1)
        '.MergeCells = True
        '.HorizontalAlignment = xlCenter
        '.VerticalAlignment = xlCenter
    End With
Next i

Application.DisplayAlerts = True

'Freeze top row
ActiveWindow.FreezePanes = False
Application.ScreenUpdating = False

Rows("2:2").Select
ActiveWindow.FreezePanes = True

ActiveWindow.FreezePanes = True
