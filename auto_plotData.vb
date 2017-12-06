'Parameter location 
'Bias H col
'Gain P col
'Pmax Q col
'Noise S col
'SNR W col
'Rise Y col
'dvdt AA col
'jitter AG col
'FWHM AK col
'Resolution AO col
'Landau AQ col

Dim ParemeterName(1 to 11) As String
ParemeterName(1) = "H"
ParemeterName(2) = "P"
ParameterName(3) = "Q"
ParameterName(4) = "S"
ParameterName(5) = "W"
ParameterName(6) = "Y"
ParameterName(7) = "AA"
ParameterName(8) = "AG"
ParameterName(9) = "AK"
ParameterName(10) = "AO"
ParameterName(11) = "AQ"


'only need the parts below

Sub CreateSheet(ByRef sheetName As String)
    Dim new_ws As Worksheet
    Set new_ws = ActiveWorkbook.Sheets.Add
    new_ws.Name = sheetName
End Sub


Sub CreateChart(ByRef PlotName As String, ByRef Yunit As String, ByRef vsWhat As String, ByRef Xunit As String, ByRef sheetLocation As String)
    Dim sh As Worksheet
    Dim chrt As Chart
    Set sh = ActiveWorkbook.Worksheets(sheetLocation)
    Set chrt = sh.Shapes.AddChart.Chart
    With chrt
        .ChartType = xlXYScatter
        .ChartArea.Border.LineStyle = xlLineStyleNone
        .Parent.Name = PlotName & " " & vsWhat
        .HasTitle = True
        .ChartTitle.Text = PlotName & " vs " & vsWhat

        .Axes(xlCategory, xlPrimary).Border.Color = vbBlack
        .Axes(xlCategory, xlPrimary).HasTitle = True
        .Axes(xlCategory, xlPrimary).HasMajorGridlines = False
        .Axes(xlCategory, xlPrimary).TickLabels.NumberFormat = "0"
        .Axes(xlCategory, xlPrimary).TickLabels.Font.Size = 12
        .Axes(xlCategory, xlPrimary).MajorTickMark = xlTickMarkInside
        .Axes(xlCategory, xlPrimary).AxisTitle.Characters.Text = vsWhat & " " & Xunit
        
        .Axes(xlValue, xlPrimary).Border.Color = vbBlack
        .Axes(xlValue, xlPrimary).HasTitle = True
        .Axes(xlValue, xlPrimary).HasMajorGridlines = False
        .Axes(xlValue, xlPrimary).TickLabels.NumberFormat = "0"
        .Axes(xlValue, xlPrimary).TickLabels.Font.Size = 12
        .Axes(xlValue, xlPrimary).MajorTickMark = xlTickMarkInside
        .Axes(xlValue, xlPrimary).AxisTitle.Characters.Text = PlotName & " " & Yunit
        '.ChartArea.AutoScaleFont = False
    End With
End Sub

Sub VsBias(ByVal sheetLoc As String, ByRef whichChart As String, ByVal index As Integer, ByVal SensorName As String, ByVal whichData As String, ByVal startRange As Integer, ByVal endRange As Integer)

    Dim sh As Worksheet
    Dim chrt As Chart
    Set sh = ActiveWorkbook.Worksheets(sheetLoc)
    Set chrt = sh.ChartObjects(whichChart).Chart
    With chrt
        .SeriesCollection.NewSeries
        .SeriesCollection(index).Name = SensorName
        .SeriesCollection(index).XValues = "'data'!$H$" & startRange & ":$H$" & endRange
        .SeriesCollection(index).Values = "'data'!$" & whichData & "$" & startRange & ":$" & whichData & "$" & endRange
    End With
End Sub

Sub VsGain(ByVal sheetLoc As String, ByRef whichChart As String, ByVal index As Integer, ByVal SensorName As String, ByVal whichData As String, ByVal startRange As Integer, ByVal endRange As Integer)

    Dim sh As Worksheet
    Dim chrt As Chart
    Set sh = ActiveWorkbook.Worksheets(sheetLoc)
    Set chrt = sh.ChartObjects(whichChart).Chart
    With chrt
        .SeriesCollection.NewSeries
        .SeriesCollection(index).Name = SensorName
        .SeriesCollection(index).XValues = "'data'!$P$" & startRange & ":$P$" & endRange
        .SeriesCollection(index).Values = "'data'!$" & whichData & "$" & startRange & ":$" & whichData & "$" & endRange
    End With
End Sub


Sub Plotting()
    
    Dim OutSheetName = "AutoPlots"

    Dim DataSetUP(4) As Variant
    DataSetUP(0) = Array("Pre-rad", 6, 10)
    DataSetUP(1) = Array("5e14 (-20C)", 28, 33)
    DataSetUP(2) = Array("5e14 (-27C)", 35, 40)
    DataSetUP(3) = Array("1e15 (-20C)", 12, 21)
    'Dim Ind As Integer

    Call CreateSheet(OutSheetName)

    Call CreateChart("Pmax", "[mV]", "Bias", "[V]", "AutoPlots")
    Call CreateChart("Gain", "", "Bias", "[V]", "AutoPlots")
    Call CreateChart("Rise Time", "[ps]", "Bias", "[V]", "AutoPlots")
    Call CreateChart("Noise", "[mV]", "Bias", "[V]", "AutoPlots")
    Call CreateChart("Jitter", "[ps]", "Bias", "[V]", "AutoPlots")
    Call CreateChart("Time Resolution", "[ps]", "Bias", "[V]", "AutoPlots")

    Call CreateChart("Pmax", "[mV]", "Gain", "", "AutoPlots")
    Call CreateChart("Rise Time", "[ps]", "Gain", "", "AutoPlots")
    Call CreateChart("Noise", "[mV]", "Gain", "", "AutoPlots")
    Call CreateChart("Jitter", "[ps]", "Gain", "", "AutoPlots")
    Call CreateChart("Time Resolution", "[ps]", "Gain", "", "AutoPlots")

        For i = 0 To 3
            Call VsBias(OutSheetName, "Pmax Bias", i, DataSetUP(i)(0), "Q", DataSetUP(i)(1), DataSetUP(i)(2))
            Call VsBias(OutSheetName, "Gain Bias", i, DataSetUP(i)(0), "P", DataSetUP(i)(1), DataSetUP(i)(2))
            Call VsBias(OutSheetName, "Rise Time Bias", i, DataSetUP(i)(0), "Y", DataSetUP(i)(1), DataSetUP(i)(2))
            Call VsBias(OutSheetName, "Noise Bias", i, DataSetUP(i)(0), "S", DataSetUP(i)(1), DataSetUP(i)(2))
            Call VsBias(OutSheetName, "Jitter Bias", i, DataSetUP(i)(0), "AG", DataSetUP(i)(1), DataSetUP(i)(2))
            Call VsBias(OutSheetName, "Time Resolution Bias", i, DataSetUP(i)(0), "AG", DataSetUP(i)(1), DataSetUP(i)(2))

            Call VsGain(OutSheetName, "Pmax Gain", i, DataSetUP(i)(0), "Q", DataSetUP(i)(1), DataSetUP(i)(2))
            Call VsGain(OutSheetName, "Rise Time Gain", i, DataSetUP(i)(0), "Y", DataSetUP(i)(1), DataSetUP(i)(2))
            Call VsGain(OutSheetName, "Noise Gain", i, DataSetUP(i)(0), "S", DataSetUP(i)(1), DataSetUP(i)(2))
            Call VsGain(OutSheetName, "Jitter Gain", i, DataSetUP(i)(0), "AG", DataSetUP(i)(1), DataSetUP(i)(2))
            Call VsGain(OutSheetName, "Time Resolution Gain", i, DataSetUP(i)(0), "AG", DataSetUP(i)(1), DataSetUP(i)(2))
        Next i
End Sub