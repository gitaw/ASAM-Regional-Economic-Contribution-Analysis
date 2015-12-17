Attribute VB_Name = "g_formatting"
'formatting====================
Sub formatManualEntry()
'lastModified 20100518
    Range("A1:I1").WrapText = False
    Range("A1:I1").HorizontalAlignment = xlLeft
    Range("A1:I1").VerticalAlignment = xlTop
End Sub
Function setRightborder(therange As Range, id)
'lastModified 20100518
    With therange.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
    Select Case id
      Case Is = 0:
        .LineStyle = xlNone
      Case Is = 1: .Weight = xlHairline
      Case Is = 2: .Weight = xlThin
      Case Is = 3: .Weight = xlMedium
    End Select
    End With
End Function

Function generalFormat()
    Range("1:1").WrapText = True
    Range("1:1").EntireRow.AutoFit
    Set myrange = Range(Cells(1, 1), Selection.SpecialCells(xlCellTypeLastCell))
    With myrange
        .Font.Name = "Calibri"
        .Font.Size = 9
        .EntireColumn.AutoFit
        .EntireRow.AutoFit
    End With
End Function
Function dollarFormat()
    Selection.Style = "Currency"
    If Not receipts = "$" Then Selection.NumberFormat = "_($* #,##0_);_($* (#,##0);_($* ""-""??_);_(@_)"
End Function
Function formatDarkmerged()
        Selection.HorizontalAlignment = xlCenter
        Selection.MergeCells = True
        Selection.Font.ThemeColor = xlThemeColorDark1
        Selection.Font.Bold = True
        Selection.Interior.Pattern = xlSolid
        Selection.Interior.ThemeColor = xlThemeColorLight2
        Selection.Borders(xlEdgeRight).LineStyle = xlContinuous
        Selection.Borders(xlEdgeRight).ThemeColor = 1
        Selection.Borders(xlEdgeRight).Weight = xlMedium
End Function


Function formatTop()
    With Selection
        .WrapText = True
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    With Selection
        .VerticalAlignment = xlTop
        .WrapText = True
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlTop
        .WrapText = True
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
        With Selection.Interior
        .PatternColorIndex = -7
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = -0.149998474074526
        .PatternTintAndShade = 0
    End With
End Function

Function freezepanes()
    For i = 2 To 10
        Sheets(mysheets(i)).Select
        With ActiveWindow
            .SplitColumn = 0
            .SplitRow = 0
        End With
        Range("b2").Select
        With ActiveWindow
            .SplitColumn = 1
            .SplitRow = 1
        End With
         ActiveWindow.freezepanes = True
    Next
End Function
Function boxin(myrange, myfill, Optional wrap)
  myrange.Borders(xlDiagonalDown).LineStyle = xlNone
    myrange.Borders(xlDiagonalUp).LineStyle = xlNone
    With myrange.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With myrange.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Weight = xlThin
    End With
    With myrange.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlThin
    End With
    With myrange.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .Weight = xlThin
    End With
    If (myfill > 0 Or myfill = True) Then
        myfill = IIf(myfill = 1 Or myfill = True, -4.99893185216834E-02, -0.149998474074526)
        With myrange.Interior
            'On Error Resume Next
            .ThemeColor = xlThemeColorDark1
            .TintAndShade = myfill
            .PatternTintAndShade = 0
        End With
    End If
    If Not IsMissing(wrap) Then
        With myrange
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = True
        End With
    End If
End Function
Sub addGotoToolsBtn(Optional nleft, Optional ntop, Optional nwidth, Optional nheight)
'kinda like a "back-button" to get to the controls
    Dim shp As Shape
    If chkShape("tools") Then _
        ActiveSheet.Shapes("tools").delete
    Set shp = ActiveSheet.Shapes.AddShape(msoShapeRoundedRectangle, _
            IIf(IsMissing(nleft), 55, nleft), IIf(IsMissing(ntop), 3, ntop), IIf(IsMissing(nwidth), 75, nwidth), IIf(IsMissing(nheight), 20, nheight))
    shp.Name = "tools"
    shp.TextFrame2.TextRange = "GoTo Tools..."
    shp.OnAction = "goTools"
    shp.Shadow.Type = msoShadow40
    shp.TextFrame2.TextRange.Font.Fill.Solid
    shp.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(0, 0, 0)
    shp.Fill.ForeColor.RGB = RGB(200, 200, 200)
    shp.Line.Weight = 0.5
    shp.Line.ForeColor.RGB = RGB(250, 250, 250)
End Sub

Function chkShape(nm)
'editing only: use when looping though named shapes
    On Error GoTo notThere
    chkShape ActiveSheet.Shapes(nm).Name
    chkShape = True
    Exit Function
notThere:
    chkShape = False
    Exit Function
End Function
'---------------------------------
'desing tool
Sub changeToolButtons()
'reset the tool button format
    Dim sr As Shape
    i = 0
    For Each sr In ActiveSheet.Shapes
        myaction = sr.OnAction
        If Not myaction = "" Then
            mytop = sr.Top
            formatButton "btn" & i, Selection.Top, Selection.OnAction
            i = i + 1
            sr.delete
        End If
    Next
End Sub
Function formatButton(nm, ntop, naction)
'consistent format for each button
    Set shp = ActiveSheet.Shapes.AddShape(msoShapeRoundedRectangle, 22, 100, 25, 12)
    shp.Name = nm
    shp.Top = ntop
    shp.left = 22
    shp.Width = 25
    shp.Height = 12
    shp.Shadow.Type = msoShadow25
    shp.Fill.ForeColor.Brightness = 0.9499999881
    shp.Line.Weight = 0.5
    shp.Line.ForeColor.RGB = RGB(250, 250, 250)
    shp.OnAction = naction
End Function


