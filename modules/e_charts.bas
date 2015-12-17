Attribute VB_Name = "e_charts"
Option Explicit
'last revised: 20151119
Public chartVariable, pieVariable, barVariable As String
Private barLimit, pieLimit, nrLines, i, j As Integer
Private id, mylbl, col1, mymax, title As Variant
Private chrt As ChartObject
'====================================
'macros called from the chart sheet buttons
Public Sub barOutput()
    updateChart "output", "bar"
End Sub
Public Sub barEmpl()
    updateChart "empl", "bar"
End Sub
Public Sub barWages()
    updateChart "wages", "bar"
End Sub
Public Sub barVa()
    updateChart "va", "bar"
End Sub
Public Sub pieOutput()
    updateChart "output", "pie"
End Sub
Public Sub pieEmpl()
    updateChart "empl", "pie"
End Sub
Public Sub pieWages()
    updateChart "wages", "pie"
End Sub
Public Sub pieVA()
    updateChart "VA", "pie"
End Sub

'-------------------------------------
' create the chart sheet(s)
Public Sub addChartSheet(chrttype, Optional parameter)
    chartVariable = IIf(IsMissing(parameter), "Output", parameter)
    id = Application.VLookup("Base" & chartVariable, ActiveWorkbook.Names("structure").RefersToRange, 11, False)
        sortTable id
    If Not WorksheetExists("chart(" & chrttype & ")") Then
        replaceSheet "chart(" & chrttype & ")"
        Sheets("chart(" & chrttype & ")").Move After:=Sheets("OutputTable")
        addNameRange chrttype & "chart_limit", Cells(1, 2)
        Cells(1, 1) = "line limit"
    'set default values for the line limit dependent on the chart type
        Cells(1, 2) = IIf(chrttype = "bar", 10, 6)
        hilite Cells(1, 2), 2
        Cells(1, 3) = "(change line limit for more or less bars and then refresh)"
        ActiveSheet.Buttons.add(5, 25, 100, 20).Select
            Selection.OnAction = chrttype & "Output"
            Selection.Characters.Text = "Output"
        ActiveSheet.Buttons.add(5, 50, 100, 20).Select
            Selection.OnAction = chrttype & "Empl"
            Selection.Characters.Text = "Employment"
        ActiveSheet.Buttons.add(5, 75, 100, 20).Select
            Selection.OnAction = chrttype & "Wages"
            Selection.Characters.Text = "Wages"
        ActiveSheet.Buttons.add(5, 100, 100, 20).Select
            Selection.OnAction = chrttype & "VA"
            Selection.Characters.Text = "Value Added (VA)"
        textbox
        ActiveSheet.Buttons.add(5, 155, 100, 20).Select
            Selection.OnAction = "exportCharts"
            Selection.Characters.Text = "Export chart(s)"
    End If
    Sheets("chart(" & chrttype & ")").Select
    If IsEmpty(ActiveWorkbook.Names(chrttype & "chart_limit")) Then ActiveWorkbook.Names(chrttype & "chart_limit").RefersToRange = IIf(chrttype = "bar", 10, 6)
    Application.Run "create" & chrttype & "chart"
End Sub

Private Sub updateChart(parameter, chrttype, Optional showmsg)
'updates existing chart to reflect different data
    Dim msg, chrt2
    'setLimit (chrttype)
    chartVariable = parameter
    Application.ScreenUpdating = False
'On Error GoTo myerror
    For i = 1 To 2
        chrt2 = chrttype
        If i = 2 Then
            chrt2 = IIf(chrttype = "bar", "pie", "bar")
            Application.ScreenUpdating = False
        End If
        Application.Run "create" & chrt2 & "chart"
    Next i
    Sheets("chart(" & chrttype & ")").Select
ending:
    Application.ScreenUpdating = True
    'If showmsg Then showinfo "Done!", True
    Exit Sub
myerror:
    GoTo ending
End Sub

Private Sub setLimit(chrttype)
'reads the desired bars/pie-wedges in the charts, or, in creating, uses the defaults
    Dim limit As Integer
    limit = ActiveWorkbook.Names(chrttype & "chart_limit").RefersToRange
    nrLines = Application.Count(Range("output[NAICS]"))
    If IsEmpty(limit) Then barLimit = IIf(chrttype = "bar", 10, 6)
'prevent error when user requests more than there are
    If limit > nrLines Then
        limit = nrLines
        Sheets("chart(" & chrttype & ")").Cells(1, 2) = nrLines
    End If
    Select Case chrttype
        Case Is = "pie": pieLimit = limit
        Case Is = "bar": barLimit = limit
    End Select
End Sub

Sub createBarchart()
'revised 20151109 -- rev4.0
' - revised title to reflect that employment numbers are not in 1,000s
' - seperated out the generic component into sub createSidebar
'revised 20111113 -- rev 3.07
' - changed series order in order to allow for series-ordered presentation in ppt
' - added a conditional in showing the $ sign in the tile so that jobs can be without
    setLimit ("bar")
    barVariable = chartVariable
    Sheets("Chart(bar)").Select
    deletecharts "chart(bar)"
    Set chrt = ActiveSheet.ChartObjects.add _
            (left:=110, Width:=475, Top:=25, Height:=425)
    chrt.Name = "Bar_" & barVariable & barLimit
    With chrt.Chart
        .charttype = xlBarStacked
        For i = 1 To 3
            If i = 2 Then mylbl = "Direct" & barVariable
            If i = 3 Then mylbl = "Indirect" & barVariable
            If i = 1 Then mylbl = "Gross" & barVariable
            id = Application.VLookup(mylbl, ActiveWorkbook.Names("structure").RefersToRange, 11, False)
            col1 = Sheets("outputTable").Rows(2).Find(id).Column
            .SeriesCollection.NewSeries
            .SeriesCollection(i).XValues = "='outputTable'!R3C3: R" & (2 + barLimit) & "C3 "
            .SeriesCollection(i).Values = "='outputTable'!R3C" & col1 & ": R" & (2 + barLimit) & "C" & col1
            If i = 2 Then .SeriesCollection(i).Name = "Direct Base " & barVariable
            If i = 3 Then .SeriesCollection(i).Name = "Indirect Base " & barVariable
            If i = 1 Then .SeriesCollection(i).Name = "Total Gross " & barVariable
        Next
        .SeriesCollection(1).AxisGroup = 2
        .SeriesCollection(1).Format.Fill.ForeColor.RGB = RGB(0, 0, 0)
        .SeriesCollection(3).Format.Fill.ForeColor.RGB = RGB(0, 51, 102)
        .SeriesCollection(2).Format.Fill.ForeColor.RGB = RGB(102, 102, 153)
        .SeriesCollection(2).Format.Fill.Transparency = 0.7
        .SeriesCollection(3).Format.Fill.Transparency = 0.7

        .ChartGroups(1).Overlap = 100
        .ChartGroups(1).GapWidth = 80
        .ChartGroups(2).Overlap = 100
        .ChartGroups(2).GapWidth = 250
        mymax = .Axes(xlValue).MaximumScale
        .Axes(xlValue, xlSecondary).MaximumScale = mymax
        .Axes(xlValue).MajorUnit = mymax / 2
        .Axes(xlValue, xlSecondary).MajorUnit = mymax / 2
        .SetElement (msoElementChartTitleAboveChart)
        title = IIf(barVariable = "empl", "Employment", IIf(barVariable = "VA", "Value Added", barVariable))
        .ChartTitle.Text = "Base vs Gross " & title & " (top " & (barLimit) & ")" & IIf(barVariable = "empl", "", " in $") & IIf(barVariable = "empl", "", Range("dollars").Value)
        .Axes(xlCategory).ReversePlotOrder = True
        .SetElement (msoElementSecondaryValueAxisNone)
        .SetElement (msoElementSecondaryCategoryAxisWithoutLabels)
        '.SetElement (msoElementSecondaryCategoryAxisShow)
        .Axes(xlCategory, xlSecondary).ReversePlotOrder = True
    End With
End Sub

Sub createPiechart()
    Sheets("Chart(pie)").Select
    deletecharts "Chart(pie)"
    setLimit ("pie")
    pieVariable = chartVariable
    For i = 1 To 2
        If i = 1 Then
            Set chrt = ActiveSheet.ChartObjects.add(left:=110, Width:=375, Top:=25, Height:=200)
            mylbl = "Base" & pieVariable
        End If
        If i = 2 Then
            Set chrt = ActiveSheet.ChartObjects.add(left:=110, Width:=375, Top:=225, Height:=200)
            mylbl = "Gross" & pieVariable
        End If
        chrt.Name = "Pie_" & mylbl & pieLimit
        id = Application.VLookup(mylbl, ActiveWorkbook.Names("structure").RefersToRange, 11, False)
        col1 = Sheets("outputTable").Rows(2).Find(id).Column
        With chrt.Chart
        'see if we have a rest-category to show in a bar next to the piechart
            .charttype = IIf(pieLimit < nrLines, xlBarOfPie, xlPie)
            'if so, calculate the desired split value
                If .charttype = xlBarOfPie Then .ChartGroups(1).SplitValue = nrLines - pieLimit
            .SeriesCollection.NewSeries
            .SeriesCollection(1).XValues = "='outputTable'!R3C3: R" & (nrLines + 2) & "C3 "
            .SeriesCollection(1).Values = "='outputTable'!R3C" & col1 & ": R" & (nrLines + 2) & "C" & col1
            .SeriesCollection(1).Name = mylbl
            .SetElement (msoElementChartTitleAboveChart)
            If i = 1 Then .ChartTitle.Text = "Base " & pieVariable
            If i = 2 Then .ChartTitle.Text = "Gross " & pieVariable
'legend
            .Legend.Font.Size = 7
            .Legend.Width = 110
            .Legend.Top = 10
            .Legend.Height = 170
            .Legend.Font.Name = "Arial"
            .Rotation = 80
            .SeriesCollection(1).ApplyDataLabels
            '.ChartGroups(1).SplitType = xlSplitByPercentValue
'show percentage for the main chart
        For j = 1 To pieLimit
            .SeriesCollection(1).Points(j).DataLabel.ShowPercentage = True
            .SeriesCollection(1).Points(j).DataLabel.ShowCategoryName = False
            .SeriesCollection(1).Points(j).DataLabel.ShowValue = False
            .SeriesCollection(1).Points(j).DataLabel.Position = xlLabelPositionBestFit
        Next
        
        For j = IIf(pieLimit < nrLines, pieLimit, 4) + 1 To IIf(pieLimit < nrLines, nrLines + 1, nrLines)
            .SeriesCollection(1).Points(j).DataLabel.ShowPercentage = False
            .SeriesCollection(1).Points(j).DataLabel.ShowValue = False
        Next
            .SeriesCollection(1).Points(nrLines + 1).DataLabel.ShowPercentage = True
            .SeriesCollection(1).Points(nrLines + 1).DataLabel.ShowCategoryName = False
            .SeriesCollection(1).Points(nrLines + 1).DataLabel.ShowValue = False
            .SeriesCollection(1).Points(nrLines + 1).DataLabel.Position = xlLabelPositionBestFit
       End With
    Next
End Sub
Public Sub exportCharts()
    j = ActiveSheet.ChartObjects.Count
    For i = 1 To j
        Set chrt = ActiveSheet.ChartObjects(i)
        chrt.Select
        ActiveChart.export Environ("USERPROFILE") & "\desktop\" & chrt.Name & ".png"
    Next i
End Sub

'support functions
        Private Sub textbox()
        'adds a textbox behind the ewxpor button to provide help
            Dim sr As Shape
            For Each sr In ActiveSheet.Shapes
             If sr.Type = msoTextBox Then _
                sr.delete
            Next sr
        ActiveSheet.Shapes.AddTextbox(msoTextOrientationHorizontal, 2, 150, 106, _
                130).Name = "exporttext"
            Set sr = ActiveSheet.Shapes("exporttext")
            sr.TextFrame.Characters.Text = _
                "Export graphs to use in reports. This export function will save the active graph to the desktop as a picture (png) file." & _
                Chr(13) & "" & Chr(13) & "NOTE: edits would normally be lost after switching output, but exporting will keep the edits."
            sr.TextFrame.Characters.Font.Size = 8
            sr.TextFrame.Characters.Font.Color = RGB(130, 130, 130)
            sr.TextFrame2.VerticalAnchor = msoAnchorBottom
            sr.Fill.ForeColor.Brightness = -0.150000006
        End Sub
        
        Sub deletecharts(where)
        On Error Resume Next
            Sheets(where).ChartObjects.delete
        End Sub
        
'=============================
Sub b_module()
'simulates the chart creation in module b
    replaceSheet ("bar")
    replaceSheet ("pie")
    addChartSheet ("bar")
    addChartSheet ("pie")
End Sub

