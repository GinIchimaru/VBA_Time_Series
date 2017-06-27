Option Explicit


Function AddChartObject1(rng As Range, Optional width As Integer = 375, Optional height As Integer = 225) As String
    Dim myChtObj As ChartObject
    Dim ChartName As String
    ChartName = rng.Address
    'if width
    
    On Error Resume Next
        ActiveSheet.ChartObjects(ChartName).Delete
    On Error GoTo 0
    
    Set myChtObj = ActiveSheet.ChartObjects.Add _
                   (Left:=100, width:=width, Top:=75, height:=height)
    
    
     '.Name = "aaaa" ActiveCell.Address
    ChartName = myChtObj.Name
    ActiveSheet.Shapes(ChartName).Name = rng.Address
    
    
    With myChtObj.Chart
        .ChartType = xlXYScatterSmoothNoMarkers
        .SetSourceData Source:=rng
        .Axes(xlCategory, xlPrimary).HasMajorGridlines = True
        .Axes(xlValue, xlPrimary).HasMajorGridlines = True
        .HasLegend = False
        .FullSeriesCollection(1).Smooth = False
        With .SeriesCollection(1).Format
            .Line.ForeColor.RGB = RGB(0, 0, 0)
            .Line.Visible = msoTrue
            .Line.Weight = 1
            With .Glow
                .Color.RGB = RGB(102, 255, 102)
                .Color.TintAndShade = 0
                .Color.Brightness = 0
                .Transparency = 0.7
                .Radius = 6
            End With
        End With
        With .Axes(xlCategory)
            With .MajorGridlines.Format.Line
                .DashStyle = msoLineDash
                .ForeColor.ObjectThemeColor = msoThemeColorText1
                .ForeColor.TintAndShade = 0
                .ForeColor.Brightness = 0
                .Transparency = 0.2
            End With
            .TickLabels.Font.Name = "Times New Roman"
            .Format.Line.ForeColor.ObjectThemeColor = msoThemeColorText1
        End With
        With .Axes(xlValue)
            With .MajorGridlines.Format.Line
                .DashStyle = msoLineDash
                .ForeColor.ObjectThemeColor = msoThemeColorText1
                .ForeColor.TintAndShade = 0
                .ForeColor.Brightness = 0
                .Transparency = 0.2
            End With
            .TickLabels.Font.Name = "Times New Roman"
            .Format.Line.ForeColor.ObjectThemeColor = msoThemeColorText1
            .Crosses = xlMinimum
        End With
    End With

    ActiveSheet.Shapes(ChartName).Line.Visible = msoFalse
    ActiveSheet.Shapes(ChartName).Fill.Visible = msoFalse
    AddChartObject1 = "Real Stats Chart " & rng.Address
    
End Function


