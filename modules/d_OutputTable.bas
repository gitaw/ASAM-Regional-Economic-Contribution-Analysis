Attribute VB_Name = "d_OutputTable"
Public hh As Collection
Sub OutputTable()
    replaceSheet ("OutputTable")
    Set myrange = Range(Cells(1, 1), Cells(je, Sheets("DataSheet").Range("1:1").Find("").Column - 1))
    myrange.FormulaR1C1 = "=datasheet!RC"
    Set myrange = Range(Cells(1, 1), Cells(je, Sheets("DataSheet").Range("1:1").Find("").Column - 1))
    addNameRange "table", myrange
    AggegateHouseholds
    ActiveSheet.ListObjects.add(xlSrcRange, myrange, , xlYes).Name = "output"
    modOutput
    sortTable
End Sub
Sub sortTable(Optional myfld)
    If IsMissing(myfld) Then myfld = "Base Output"
    Worksheets("OutputTable").ListObjects("output").Sort.SortFields.Clear
    Worksheets("OutputTable").ListObjects("output").Sort.SortFields.add key:= _
        Range("output[[#All],[" & myfld & "]]"), _
        SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("OutputTable").ListObjects("output").Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
End Sub
Function modOutput()
    Dim structure As Variant
    Set myrange = ActiveWorkbook.Names("structure").RefersToRange
    With myrange.CurrentRegion
        r = .Rows.Count
        c = .Columns.Count
        structure = .Resize(r, c)
        Sheets("outputTable").Select
        For i = 1 To r - 1
            Cells(1, i).FormulaR1C1 = structure(i + 1, 11)
            ActiveSheet.Columns(i).Select
            If Not structure(i + 1, 6) = "na" Then ActiveSheet.Columns(i).NumberFormat = structure(i + 1, 6)
        Next i
        ActiveSheet.Rows(1).Insert Shift:=xlDown
        generalFormat
        mergcnt = 0
        For i = 1 To r - 1
            ActiveSheet.Columns(i).EntireColumn.Hidden = structure(i + 1, 8)
            If Not Nz(structure(i + 1, 9), "") = "" Then
                Cells(1, i).FormulaR1C1 = structure(i + 1, 9)
                mergecnt = 0
            End If
            If mergecnt > 1 Then mergeLabel Range(Cells(1, i - mergecnt), Cells(1, i))
            mergecnt = mergecnt + 1
        Next i
    End With
    Sheets("outputTable").Cells(3, 4).Select
    ActiveWindow.freezepanes = True
End Function

Sub mergeLabel(thisrange)
    boxin thisrange, 2
        With thisrange
            .Select
            .MergeCells = True
            .Font.Size = 12
            .Font.Bold = True
            .HorizontalAlignment = xlLeft
            .VerticalAlignment = xlCenter
        End With
End Sub
Function HHoccurance()
'lastModified 20111113 rev 3.1
' -- there is a 814 private households NAICS sector so I added a conditional to look for the 814 and not count it
    HHoccurance = 0
    Set hh = New Collection
    For i = 2 To je
        On Error Resume Next
        myval = Sheets("datasheet").Cells(i, 3).Find(what:="Households", LookIn:=xlValues, _
            LookAt:=xlPart, MatchCase:=False, SearchFormat:=False)
        If err.Number = 0 Then
            If InStr(1, myval, "814", vbTextCompare) = 0 Then
                HHoccurance = HHoccurance + 1
                hh.add i
            End If
        End If
    Next
End Function

Sub AggegateHouseholds()
    mergeHouseholds = ActiveWorkbook.Names("mergehouseholds").RefersToRange
'lastModified 20100518
    Sheets("outputtable").Select
    n_hh = HHoccurance
    If n_hh < 2 Then Exit Sub
    Select Case mergeHouseholds
        Case Is = 1: 'merge into three tiers
            If n_hh < 7 Then
                MsgBox ("There are only " & n_hh & " household groups; Households will be aggregated into one group...")
                GoTo HHreroute
            End If
HHreroute:
            Cells(hh(1), 3).Value = "Households (aggregate)"
            Range(Cells(hh(1), 4), Cells(hh(1), Rows(hh(1)).Find("").Column - 1)).FormulaR1C1 = _
                 "=sum(datasheet!R" & hh(1) & "C:R" & hh(hh.Count) & "C)"
            mystr = hh(2) & ":" & hh(hh.Count)
            Rows(mystr).delete Shift:=xlUp
            Exit Sub
        Case Is = 2: 'aggregate into three groups
            Cells(hh(1), 3).Value = "Households (low tier)"
            Range(Cells(hh(1), 4), Cells(hh(1), Rows(hh(1)).Find("").Column - 1)).FormulaR1C1 = _
                     "=sum(datasheet!R" & hh(1) & "C:R" & hh(3) & "C)"
            Cells(hh(4), 3).Value = "Households (middle tier)"
            Range(Cells(hh(4), 4), Cells(hh(4), Rows(hh(4)).Find("").Column - 1)).Select
            Range(Cells(hh(4), 4), Cells(hh(4), Rows(hh(4)).Find("").Column - 1)).FormulaR1C1 = _
                     "=sum(datasheet!R" & hh(4) & "C:R" & hh(6) & "C)"
            Cells(hh(7), 3).Value = "Households (high tier)"
            Range(Cells(hh(7), 4), Cells(hh(7), Rows(hh(7)).Find("").Column - 1)).Select
            Range(Cells(hh(7), 4), Cells(hh(7), Rows(hh(7)).Find("").Column - 1)).FormulaR1C1 = _
                     "=sum(datasheet!R" & hh(7) & "C:R" & hh(hh.Count) & "C)"
            Rows(hh(8) & ":" & hh(hh.Count)).delete Shift:=xlUp
            Rows(hh(5) & ":" & hh(6)).delete Shift:=xlUp
            Rows(hh(2) & ":" & hh(3)).delete Shift:=xlUp
        Case Is = 3: Exit Sub 'no aggregation
    End Select
End Sub
