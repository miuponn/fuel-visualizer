ActiveSheet.Shapes.AddChart2(297, xlColumnStacked).Select
    ActiveChart.SetSourceData Source:=Range( _
        "'2022 Output By Fuel Type'!$A$1:$N$7")
    ActiveSheet.Shapes("Chart 3").ScaleWidth 1.6789930009, msoFalse, _
        msoScaleFromTopLeft
    ActiveSheet.Shapes("Chart 3").ScaleHeight 1.8184971098, msoFalse, _
        msoScaleFromTopLeft
    ActiveChart.Legend.Select
    ActiveChart.Legend.LegendEntries(1).Select
    With Selection.Format.Fill
        .Visible = msoTrue
        .ForeColor.ObjectThemeColor = msoThemeColorAccent6
        .ForeColor.TintAndShade = 0
        .ForeColor.Brightness = 0
        .Transparency = 0
        .Solid
    End With
    ActiveChart.Legend.LegendEntries(2).Select
    ActiveChart.Legend.LegendEntries(3).Select
    With Selection.Format.Fill
        .Visible = msoTrue
        .ForeColor.ObjectThemeColor = msoThemeColorAccent5
        .ForeColor.TintAndShade = 0
        .ForeColor.Brightness = 0
        .Transparency = 0
        .Solid
    End With
    ActiveChart.Legend.LegendEntries(4).Select
    With Selection.Format.Fill
        .Visible = msoTrue
        .ForeColor.ObjectThemeColor = msoThemeColorAccent3
        .ForeColor.TintAndShade = 0
        .ForeColor.Brightness = 0
        .Transparency = 0
        .Solid
    End With
    ActiveChart.Legend.LegendEntries(5).Select
    With Selection.Format.Fill
        .Visible = msoTrue
        .ForeColor.ObjectThemeColor = msoThemeColorAccent4
        .ForeColor.TintAndShade = 0
        .ForeColor.Brightness = 0
        .Transparency = 0
        .Solid
    End With
    ActiveChart.Legend.LegendEntries(6).Select
    With Selection.Format.Fill
        .Visible = msoTrue
        .ForeColor.ObjectThemeColor = msoThemeColorAccent5
        .ForeColor.TintAndShade = 0
        .ForeColor.Brightness = 0.8000000119
        .Transparency = 0
        .Solid
    End With
    ActiveChart.SetElement (msoElementPrimaryCategoryAxisTitleAdjacentToAxis)
    ActiveChart.SetElement (msoElementPrimaryValueAxisTitleAdjacentToAxis)
    ActiveChart.Axes(xlCategory).AxisTitle.Select
    Selection.Delete
    ActiveSheet.ChartObjects("Chart 3").Activate
    ActiveChart.Axes(xlValue).AxisTitle.Select
    ActiveChart.Axes(xlValue, xlPrimary).AxisTitle.Text = "Monthly Output (MW)"
    Selection.Format.TextFrame2.TextRange.Characters.Text = "Monthly Output (MW)"
    With Selection.Format.TextFrame2.TextRange.Characters(1, 19).ParagraphFormat
        .TextDirection = msoTextDirectionLeftToRight
        .Alignment = msoAlignCenter
    End With
    With Selection.Format.TextFrame2.TextRange.Characters(1, 7).Font
        .BaselineOffset = 0
        .Bold = msoFalse
        .NameComplexScript = "+mn-cs"
        .NameFarEast = "+mn-ea"
        .Fill.Visible = msoTrue
        .Fill.ForeColor.RGB = RGB(89, 89, 89)
        .Fill.Transparency = 0
        .Fill.Solid
        .Size = 10
        .Italic = msoFalse
        .Kerning = 12
        .name = "+mn-lt"
        .UnderlineStyle = msoNoUnderline
        .Strike = msoNoStrike
    End With
    With Selection.Format.TextFrame2.TextRange.Characters(8, 12).Font
        .BaselineOffset = 0
        .Bold = msoFalse
        .NameComplexScript = "+mn-cs"
        .NameFarEast = "+mn-ea"
        .Fill.Visible = msoTrue
        .Fill.ForeColor.RGB = RGB(89, 89, 89)
        .Fill.Transparency = 0
        .Fill.Solid
        .Size = 10
        .Italic = msoFalse
        .Kerning = 12
        .name = "+mn-lt"
        .UnderlineStyle = msoNoUnderline
        .Strike = msoNoStrike
    End With
    ActiveChart.ChartTitle.Select
    ActiveChart.ChartTitle.Text = "Monthly Output By Fuel Type"
    Selection.Format.TextFrame2.TextRange.Characters.Text = _
        "Monthly Output By Fuel Type"
    With Selection.Format.TextFrame2.TextRange.Characters(1, 27).ParagraphFormat
        .TextDirection = msoTextDirectionLeftToRight
        .Alignment = msoAlignCenter
    End With
    With Selection.Format.TextFrame2.TextRange.Characters(1, 7).Font
        .BaselineOffset = 0
        .Bold = msoFalse
        .NameComplexScript = "+mn-cs"
        .NameFarEast = "+mn-ea"
        .Fill.Visible = msoTrue
        .Fill.ForeColor.RGB = RGB(89, 89, 89)
        .Fill.Transparency = 0
        .Fill.Solid
        .Size = 14
        .Italic = msoFalse
        .Kerning = 12
        .name = "+mn-lt"
        .UnderlineStyle = msoNoUnderline
        .Spacing = 0
        .Strike = msoNoStrike
    End With
    With Selection.Format.TextFrame2.TextRange.Characters(8, 20).Font
        .BaselineOffset = 0
        .Bold = msoFalse
        .NameComplexScript = "+mn-cs"
        .NameFarEast = "+mn-ea"
        .Fill.Visible = msoTrue
        .Fill.ForeColor.RGB = RGB(89, 89, 89)
        .Fill.Transparency = 0
        .Fill.Solid
        .Size = 14
        .Italic = msoFalse
        .Kerning = 12
        .name = "+mn-lt"
        .UnderlineStyle = msoNoUnderline
        .Spacing = 0
        .Strike = msoNoStrike
    End With

'Copy paste chart onto new sheet
ws.ChartObjects("Chart ").CopyPicture xlScreen, xlPicture
Worksheets("Destination").Select
Cells(pasterow, 1).Select
ActiveSheet.Paste

'Create stacked chart
Dim rng5 As Range
Dim MyChart As Chart

Set MyChart = Charts.Add

With MyChart
    .ChartType = xlColumnStacked
    .SetSourceData rng5
    .ChartTitle.Text = "Monthly Output By Fuel Type"
    .HasLegend = True
    .SeriesCollection(1).Interior.Color = RGB(112, 173, 71) 'green biofuel
    .SeriesCollection(2).Interior.Color = RGB(165, 165, 165) 'grey gas
    .SeriesCollection(3).Interior.Color = RGB(91, 155, 213) 'blue hydro
    .SeriesCollection(4).Interior.Color = RGB(0, 43, 130) 'indigo nuclear
    .SeriesCollection(5).Interior.Color = RGB(255, 192, 0) 'yellow solar
    .SeriesCollection(6).Interior.Color = RGB(222, 235, 247) 'light blue wind
    .Axes(xlValue, xlPrimary).HasTitle = True
    .Axes(xlValue, xlPrimary).AxisTitle.Characters.Text = "Monthly Output (MW)"
End With

ws.Cells(1, 16).Value = "Output values displayed in MW"

End Sub
