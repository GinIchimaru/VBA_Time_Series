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
    myChtObj.Chart.SetSourceData Source:=rng
    myChtObj.Chart.ChartType = xlXYScatterSmoothNoMarkers
     '.Name = "aaaa" ActiveCell.Address
    ChartName = myChtObj.Name
    ActiveSheet.Shapes(ChartName).Name = rng.Address
    
    
    myChtObj.Chart.HasLegend = False
    myChtObj.Chart.FullSeriesCollection(1).Smooth = False
    
    With myChtObj.Chart
    .Axes(xlCategory, xlPrimary).HasMajorGridlines = True
    .Axes(xlValue, xlPrimary).HasMajorGridlines = True
    End With

    With myChtObj.Chart.SeriesCollection(1).Format
        .Line.ForeColor.RGB = RGB(0, 0, 0)
        .Line.Visible = msoTrue
        .Line.Weight = 1
        With .Glow
            .Color.RGB = RGB(102, 255, 102)
            .Color.TintAndShade = 0
            .Color.Brightness = 0
            .Transparency = 0.8000000119
            .Radius = 6
        End With
    
    End With
    
   With myChtObj.Chart.Axes(xlCategory).MajorGridlines.Format.Line
        .DashStyle = msoLineDash
        .ForeColor.ObjectThemeColor = msoThemeColorText1
        .ForeColor.TintAndShade = 0
        .ForeColor.Brightness = 0
        .Transparency = 0.2
    End With
    With myChtObj.Chart.Axes(xlValue).MajorGridlines.Format.Line
        .DashStyle = msoLineDash
        .ForeColor.ObjectThemeColor = msoThemeColorText1
        .ForeColor.TintAndShade = 0
        .ForeColor.Brightness = 0
        .Transparency = 0.2
    End With
    ActiveSheet.Shapes(ChartName).Line.Visible = msoFalse
    ActiveSheet.Shapes(ChartName).Fill.Visible = msoFalse


    AddChartObject1 = "Real Stats Chart " & rng.Address
    
End Function


Sub Macro25()
'
' Macro25 Macro
'

'
    ActiveSheet.ChartObjects("$A$1:$A$200").Activate
    ActiveChart.Parent.Delete
End Sub
Sub Macro26()
'
' Macro26 Macro
'

'
    ActiveSheet.ChartObjects("$A$1:$A$200").Activate
    ActiveSheet.ChartObjects("$A$1:$A$200").Activate
    ActiveSheet.Shapes("$A$1:$A$200").Fill.Visible = msoFalse
End Sub
Sub Macro27()
'
' Macro27 Macro
'

'
    ActiveSheet.Shapes("$A$1:$A$200").Line.Visible = msoFalse
End Sub