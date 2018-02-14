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

'Dim ParemeterName(1 to 11) As String
'ParemeterName(1) = "H"
'ParemeterName(2) = "P"
'ParameterName(3) = "Q"
'ParameterName(4) = "S"
'ParameterName(5) = "W"
'ParameterName(6) = "Y"
'ParameterName(7) = "AA"
'ParameterName(8) = "AG"
'ParameterName(9) = "AK"
'ParameterName(10) = "AO"
'ParameterName(11) = "AQ"


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
        .ChartTitle.Font.Size = 12

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


Sub VsGain(ByVal setToSheetLoc As String, ByRef whichChart As String, ByVal index As Integer, ByVal whichPage, ByVal SensorNameCol As String, ByVal SensorNameRow As Integer, ByVal whichData As String, ByVal startRange As Integer, ByVal endRange As Integer, ByVal symbolShape As String, ByVal symbolColor As String, ByVal fillOrNofill As String)

    Dim dataPtShape As String
    dataPtShape = symbolShape

    Dim sh As Worksheet
    Dim chrt As Chart
    Set sh = ActiveWorkbook.Worksheets(setToSheetLoc)
    Set chrt = sh.ChartObjects(whichChart).Chart
    With chrt
        .SeriesCollection.NewSeries
        .SeriesCollection(index).Name = "=" & whichPage & "!$" & SensorNameCol &  "$" & SensorNameRow
        .SeriesCollection(index).XValues = whichPage & "!$P$" & startRange & ":$P$" & endRange
        .SeriesCollection(index).Values = whichPage & "!$" & whichData & "$" & startRange & ":$" & whichData & "$" & endRange
        
        Select Case dataPtShape
            Case Is = "square"
                .SeriesCollection(index).MarkerStyle = xlMarkerStyleSquare
            Case Is = "circle"
                .SeriesCollection(index).MarkerStyle = xlMarkerStyleCircle
            Case Is = "diamond"
                .SeriesCollection(index).MarkerStyle = xlMarkerStyleDiamond 
            Case Is = "star"
                .SeriesCollection(index).MarkerStyle = xlMarkerStyleStar 
            Case Is = "triangle"
                .SeriesCollection(index).MarkerStyle = xlMarkerStyleTriangle 
            Case Is = "x"
                .SeriesCollection(index).MarkerStyle = xlMarkerStyleX 
        End Select

        Select Case fillOrNofill
            Case Is = "unfill"
                .SeriesCollection(index).Format.Fill.Visible = msoFalse
                With .SeriesCollection(index).Format.Line
                    .Visible = msoFalse
                    .Transparency = 0
                End With

                With .SeriesCollection(index)
                    Select Case symbolColor
                        Case Is = "red"
                            .MarkerForegroundColor = RGB(255, 0, 0)
                        Case Is = "yellow"
                            .MarkerForegroundColor = RGB(255, 192, 0)
                        Case Is = "green"
                            .MarkerForegroundColor = RGB(0, 176, 80)
                        Case Is = "black"
                            .MarkerForegroundColor = vbBlack
                        Case Is = "blue"
                            .MarkerForegroundColor = vbBlue
                        Case Is = "magenta"
                            .MarkerForegroundColor = vbMagenta
                        Case Is = "cyan"
                            .MarkerForegroundColor = vbCyan
                        Case Is = "purple"
                            .MarkerForegroundColor = RGB(255, 0, 255)
                    End Select
                End With

            Case Is = "fill"
                With .SeriesCollection(index).Format.Fill
                    .Visible = msoTrue
                    .Solid
                    .Transparency = 0
                    Select Case symbolColor
                        Case Is = "red"
                            .ForeColor.RGB = RGB(255, 0, 0)
                        Case Is = "yellow"
                            .ForeColor.RGB = RGB(255, 192, 0)
                        Case Is = "green"
                            .ForeColor.RGB = RGB(0, 176, 80)
                        Case Is = "black"
                            .ForeColor.RGB = vbBlack
                        Case Is = "blue"
                            .ForeColor.RGB = vbBlue
                        Case Is = "magenta"
                            .ForeColor.RGB = vbMagenta
                        Case Is = "cyan"
                            .ForeColor.RGB = vbCyan
                        Case Is = "purple"
                            .ForeColor.RGB = RGB(255, 0, 255)
                     End Select
                End With

                With .SeriesCollection(index).Format.Line
                    .Visible = msoFalse
                    .Transparency = 0
                End With

                With .SeriesCollection(index)
                    Select Case symbolColor
                        Case Is = "red"
                            .MarkerForegroundColor = RGB(255, 0, 0)
                        Case Is = "yellow"
                            .MarkerForegroundColor = RGB(255, 192, 0)
                        Case Is = "green"
                            .MarkerForegroundColor = RGB(0, 176, 80)
                        Case Is = "black"
                            .MarkerForegroundColor = vbBlack
                        Case Is = "blue"
                            .MarkerForegroundColor = vbBlue
                        Case Is = "magenta"
                            .MarkerForegroundColor = vbMagenta
                        Case Is = "cyan"
                            .MarkerForegroundColor = vbCyan
                        Case Is = "purple"
                            .MarkerForegroundColor = RGB(255, 0, 255)
                    End Select
                End With
        End Select
    End With
End Sub

Sub VsBias(ByVal setToSheetLoc As String, ByRef whichChart As String, ByVal index As Integer, ByVal whichPage, ByVal SensorNameCol As String, ByVal SensorNameRow As Integer, ByVal whichData As String, ByVal startRange As Integer, ByVal endRange As Integer, ByVal symbolShape As String, ByVal symbolColor As String, ByVal fillOrNofill As String)

    Dim dataPtShape As String
    dataPtShape = symbolShape

    Dim sh As Worksheet
    Dim chrt As Chart
    Set sh = ActiveWorkbook.Worksheets(setToSheetLoc)
    Set chrt = sh.ChartObjects(whichChart).Chart
    With chrt
        .SeriesCollection.NewSeries
        .SeriesCollection(index).Name = "=" & whichPage & "!$" & SensorNameCol &  "$" & SensorNameRow
        .SeriesCollection(index).XValues = whichPage & "!$H$" & startRange & ":$H$" & endRange
        .SeriesCollection(index).Values = whichPage & "!$" & whichData & "$" & startRange & ":$" & whichData & "$" & endRange
        
        Select Case dataPtShape
            Case Is = "square"
                .SeriesCollection(index).MarkerStyle = xlMarkerStyleSquare
            Case Is = "circle"
                .SeriesCollection(index).MarkerStyle = xlMarkerStyleCircle
            Case Is = "diamond"
                .SeriesCollection(index).MarkerStyle = xlMarkerStyleDiamond 
            Case Is = "star"
                .SeriesCollection(index).MarkerStyle = xlMarkerStyleStar 
            Case Is = "triangle"
                .SeriesCollection(index).MarkerStyle = xlMarkerStyleTriangle 
            Case Is = "x"
                .SeriesCollection(index).MarkerStyle = xlMarkerStyleX 
        End Select
        
        Select Case fillOrNofill
            Case Is = "unfill"
                .SeriesCollection(index).Format.Fill.Visible = msoFalse
                With .SeriesCollection(index).Format.Line
                    .Visible = msoFalse
                    .Transparency = 0
                End With

                With .SeriesCollection(index)
                    Select Case symbolColor
                        Case Is = "red"
                            .MarkerForegroundColor = RGB(255, 0, 0)
                        Case Is = "yellow"
                            .MarkerForegroundColor = RGB(255, 192, 0)
                        Case Is = "green"
                            .MarkerForegroundColor = RGB(0, 176, 80)
                        Case Is = "black"
                            .MarkerForegroundColor = vbBlack
                        Case Is = "blue"
                            .MarkerForegroundColor = vbBlue
                        Case Is = "magenta"
                            .MarkerForegroundColor = vbMagenta
                        Case Is = "cyan"
                            .MarkerForegroundColor = vbCyan
                        Case Is = "purple"
                            .MarkerForegroundColor = RGB(255, 0, 255)
                    End Select
                End With

            Case Is = "fill"
                With .SeriesCollection(index).Format.Fill
                    .Visible = msoTrue
                    .Solid
                    .Transparency = 0
                    Select Case symbolColor
                        Case Is = "red"
                            .ForeColor.RGB = vbRed
                        Case Is = "yellow"
                            .ForeColor.RGB = vbYellow
                        Case Is = "green"
                            .ForeColor.RGB = vbGreen
                        Case Is = "black"
                            .ForeColor.RGB = vbBlack
                        Case Is = "blue"
                            .ForeColor.RGB = vbBlue
                        Case Is = "magenta"
                            .ForeColor.RGB = vbMagenta
                        Case Is = "cyan"
                            .ForeColor.RGB = vbCyan
                        Case Is = "purple"
                            .ForeColor.RGB = vbPurple
                     End Select
                End With

                With .SeriesCollection(index).Format.Line
                    .Visible = msoFalse
                    .Transparency = 0
                End With

                With .SeriesCollection(index)
                    Select Case symbolColor
                        Case Is = "red"
                            .MarkerForegroundColor = RGB(255, 0, 0)
                        Case Is = "yellow"
                            .MarkerForegroundColor = RGB(255, 192, 0)
                        Case Is = "green"
                            .MarkerForegroundColor = RGB(0, 176, 80)
                        Case Is = "black"
                            .MarkerForegroundColor = vbBlack
                        Case Is = "blue"
                            .MarkerForegroundColor = vbBlue
                        Case Is = "magenta"
                            .MarkerForegroundColor = vbMagenta
                        Case Is = "cyan"
                            .MarkerForegroundColor = vbCyan
                        Case Is = "purple"
                            .MarkerForegroundColor = RGB(255, 0, 255)
                    End Select
                End With
        End Select
    End With
End Sub

Sub Plotting()
    
    Dim OutSheetName as String
    OutSheetName = "HPK50D W6W8 plots v2"

    Dim DataSetUP(11) As Variant
    DataSetUP(0) = Array("A", 9, 14)

    DataSetUP(1) = Array("W6vsW8", "B", 55, 56, 58, "diamond", "green", "unfill")
    DataSetUP(2) = Array("W6vsW8", "B", 8, 9, 14, "diamond", "yellow", "unfill")
    DataSetUP(3) = Array("W6vsW8", "B", 32, 33, 38, "diamond", "purple", "unfill")

    DataSetUP(4) = Array("W6vsW8", "B", 60, 61, 63, "square", "green", "unfill")
    DataSetUP(5) = Array("W6vsW8", "B", 18, 19, 24, "square", "yellow", "unfill")
    DataSetUP(6) = Array("W6vsW8", "B", 40, 41, 44, "square", "purple", "unfill")

    DataSetUP(7) = Array("'HPK 50D data'", "B", 29, 30, 35, "circle", "green", "unfill")
    DataSetUP(8) = Array("'HPK 50D data'", "B", 38, 39, 46, "circle", "yellow", "unfill")
    DataSetUP(9) = Array("'HPK 50D data'", "B", 49, 50, 56, "circle", "red", "unfill")
    DataSetUP(10) = Array("'HPK 50D data'", "B", 59, 60, 66, "circle", "purple", "unfill")
    'DataSetUP(2) = Array("C", 19, 24, "circle", "blue", "unfill")
    'DataSetUP(3) = Array("D", 41, 44, "x", "yellow", "fill")
    'Dim Ind As Integer

    Call CreateSheet(OutSheetName)

    Call CreateChart("Pmax", "[mV]", "Bias", "[V]", OutSheetName)
    Call CreateChart("Gain", "", "Bias", "[V]", OutSheetName)
    Call CreateChart("Rise Time", "[ps]", "Bias", "[V]", OutSheetName)
    Call CreateChart("Noise", "[mV]", "Bias", "[V]", OutSheetName)
    Call CreateChart("Jitter CFD20", "[ps]", "Bias", "[V]", OutSheetName)
    Call CreateChart("Time Resolution CFD20", "[ps]", "Bias", "[V]", OutSheetName)

    Call CreateChart("Pmax", "[mV]", "Gain", "", OutSheetName)
    Call CreateChart("Rise Time", "[ps]", "Gain", "", OutSheetName)
    Call CreateChart("Noise", "[mV]", "Gain", "", OutSheetName)
    Call CreateChart("Jitter CFD20", "[ps]", "Gain", "", OutSheetName)
    Call CreateChart("Time Resolution CFD20", "[ps]", "Gain", "", OutSheetName)

    Dim whichPageReferTo As String
    whichPageReferTo = "W6vsW8"

    Dim whichPageReferTo2 As String
    whichPageReferTo2 = "HPK 50D data"
 
 'VsBias(ByVal setToSheetLoc As String, ByRef whichChart As String, ByVal index As Integer, ByVal whichPage, ByVal SensorNameCol As String, ByVal SensorNameRow As Integer, ByVal whichData As String, ByVal startRange As Integer, ByVal endRange As Integer, ByVal symbolShape As String, ByVal symbolColor As String, ByVal fillOrNofill As String)
 
        For i = 1 To 10
            Call VsBias(OutSheetName, "Gain Bias", i, DataSetUP(i)(0), DataSetUP(i)(1), DataSetUP(i)(2), "P", DataSetUP(i)(3), DataSetUP(i)(4),  DataSetUP(i)(5),  DataSetUP(i)(6),  DataSetUP(i)(7))

            Call VsBias(OutSheetName, "Pmax Bias", i, DataSetUP(i)(0), DataSetUP(i)(1), DataSetUP(i)(2), "Q", DataSetUP(i)(3), DataSetUP(i)(4),  DataSetUP(i)(5),  DataSetUP(i)(6),  DataSetUP(i)(7))

            Call VsBias(OutSheetName, "Noise Bias", i, DataSetUP(i)(0), DataSetUP(i)(1), DataSetUP(i)(2), "S", DataSetUP(i)(3), DataSetUP(i)(4),  DataSetUP(i)(5),  DataSetUP(i)(6),  DataSetUP(i)(7))

            Call VsBias(OutSheetName, "Rise Time Bias", i, DataSetUP(i)(0), DataSetUP(i)(1), DataSetUP(i)(2), "Y", DataSetUP(i)(3), DataSetUP(i)(4),  DataSetUP(i)(5),  DataSetUP(i)(6),  DataSetUP(i)(7))

            If i > 6 Then
            Call VsBias(OutSheetName, "Jitter CFD20 Bias", i, DataSetUP(i)(0), DataSetUP(i)(1), DataSetUP(i)(2), "AG", DataSetUP(i)(3), DataSetUP(i)(4),  DataSetUP(i)(5),  DataSetUP(i)(6),  DataSetUP(i)(7))

            Call VsBias(OutSheetName, "Time Resolution CFD20 Bias", i, DataSetUP(i)(0), DataSetUP(i)(1), DataSetUP(i)(2), "AO", DataSetUP(i)(3), DataSetUP(i)(4),  DataSetUP(i)(5),  DataSetUP(i)(6),  DataSetUP(i)(7))
            Else
            Call VsBias(OutSheetName, "Jitter CFD20 Bias", i, DataSetUP(i)(0), DataSetUP(i)(1), DataSetUP(i)(2), "AI", DataSetUP(i)(3), DataSetUP(i)(4),  DataSetUP(i)(5),  DataSetUP(i)(6),  DataSetUP(i)(7))

            Call VsBias(OutSheetName, "Time Resolution CFD20 Bias", i, DataSetUP(i)(0), DataSetUP(i)(1), DataSetUP(i)(2), "AQ", DataSetUP(i)(3), DataSetUP(i)(4),  DataSetUP(i)(5),  DataSetUP(i)(6),  DataSetUP(i)(7))
            End IF
            
            Call VsGain(OutSheetName, "Pmax Gain", i, DataSetUP(i)(0), DataSetUP(i)(1), DataSetUP(i)(2), "Q", DataSetUP(i)(3), DataSetUP(i)(4),  DataSetUP(i)(5),  DataSetUP(i)(6),  DataSetUP(i)(7))

            Call VsGain(OutSheetName, "Noise Gain", i, DataSetUP(i)(0), DataSetUP(i)(1), DataSetUP(i)(2), "S", DataSetUP(i)(3), DataSetUP(i)(4),  DataSetUP(i)(5),  DataSetUP(i)(6),  DataSetUP(i)(7))

            Call VsGain(OutSheetName, "Rise Time Gain", i, DataSetUP(i)(0), DataSetUP(i)(1), DataSetUP(i)(2), "Y", DataSetUP(i)(3), DataSetUP(i)(4),  DataSetUP(i)(5),  DataSetUP(i)(6),  DataSetUP(i)(7))

            If i > 6 Then
            Call VsGain(OutSheetName, "Jitter CFD20 Gain", i, DataSetUP(i)(0), DataSetUP(i)(1), DataSetUP(i)(2), "AG", DataSetUP(i)(3), DataSetUP(i)(4),  DataSetUP(i)(5),  DataSetUP(i)(6),  DataSetUP(i)(7))

            Call VsGain(OutSheetName, "Time Resolution CFD20 Gain", i, DataSetUP(i)(0), DataSetUP(i)(1), DataSetUP(i)(2), "AO", DataSetUP(i)(3), DataSetUP(i)(4),  DataSetUP(i)(5),  DataSetUP(i)(6),  DataSetUP(i)(7))
            Else
            Call VsGain(OutSheetName, "Jitter CFD20 Gain", i, DataSetUP(i)(0), DataSetUP(i)(1), DataSetUP(i)(2), "AI", DataSetUP(i)(3), DataSetUP(i)(4),  DataSetUP(i)(5),  DataSetUP(i)(6),  DataSetUP(i)(7))

            Call VsGain(OutSheetName, "Time Resolution CFD20 Gain", i, DataSetUP(i)(0), DataSetUP(i)(1), DataSetUP(i)(2), "AQ", DataSetUP(i)(3), DataSetUP(i)(4),  DataSetUP(i)(5),  DataSetUP(i)(6),  DataSetUP(i)(7))
            End IF
        Next i
End Sub