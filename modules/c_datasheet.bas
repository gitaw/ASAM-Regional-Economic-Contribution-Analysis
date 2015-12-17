Attribute VB_Name = "c_datasheet"
Public coll_tag As Collection
Public coll_memo As Collection

Sub initDatasheet()
    Dim structure As Variant
Set coll_tag = New Collection
    replaceSheet ("DataSheet")
'read the table structure into tha array
    Sheets("structure").Visible = True
    Sheets("structure").Select
    Set myrange = Sheets("structure").Range(Cells(1, 1), Cells(Range("K:K").Find("").Row - 1, 11))
    addNameRange "structure", myrange
    With myrange.CurrentRegion
        r = .Rows.Count
        c = .Columns.Count
        structure = .Resize(r, c)
        Sheets("DataSheet").Select
        For i = 1 To r - 1
            Cells(1, i).FormulaR1C1 = structure(i + 1, 3) & Chr(10) & structure(i + 1, 4)
            Cells(js + 2, i) = i
            If Not Nz(structure(i + 1, 10), "") = "" Then
                Cells(1, i).AddComment
                Cells(1, i).Comment.Visible = False
                Cells(1, i).Comment.Text Text:=structure(i + 1, 10)
            End If
            If i > 1 Then coll_tag.add structure(i, 1)
        Next i
        coll_tag.add structure(i, 1)
        Rows(js + 2).EntireRow.Hidden = True
    End With
    Set myrange = Range(Cells(1, 1), Cells(1, r - 1))
    boxin myrange, True, True
    fillinDatasheet ("group")
'create totals line
    Cells(js + 1, 3) = "Totals"
    Range(Cells(js + 1, 4), Cells(js + 1, r - 1)) = "=sum(R2C:R" & js & "C)"
'---add the columns that are needed to create the matrices
'   we'll do the rest later
'add employment
    fillinDatasheet ("grossEmpl")
'add wages
    fillinDatasheet ("GrossWages")
'add Gross valueAdded
    fillinDatasheet ("GrossVA")
'add endo/exo
    fillinDatasheet ("Exogenous")
    fillinDatasheet ("Endogenous")
End Sub
Sub finishDatasheet(Optional msg)
    Calculate
    If Not IsMissing(msg) Then
        msg = msg & " - calculating..." & vbCrLf
        showinfo msg
    End If
    For i = 1 To coll_tag.Count
        fillinDatasheet coll_tag(i)
    Next
    Sheets("structure").Visible = False
    generalFormat
    datasheetformat
'add a navigation button
'   first make space
    Range("c1") = Application.WorksheetFunction.clean(Range("c1")) 'take out carriage returns
    Range("c1").HorizontalAlignment = xlLeft
    Range("c1").VerticalAlignment = xlBottom
    Range("c1").IndentLevel = 1
    addGotoToolsBtn
End Sub

Sub fillinDatasheet(id)
    Sheets("DataSheet").Select
    mycol = Application.VLookup(id, ActiveWorkbook.Names("structure").RefersToRange, 2, False)
    myformat = Application.VLookup(id, ActiveWorkbook.Names("structure").RefersToRange, 6, False)
    Set myrange = Range(Cells(1, mycol), Cells(js + 1, mycol))
    If Not myformat = "na" Then myrange.NumberFormat = myformat
    setRightborder myrange, Application.VLookup(id, ActiveWorkbook.Names("structure").RefersToRange, 7, False)
Select Case id
Case Is = "group":
    Sheets("DataSheet").Cells(jc, 1) = "end of sectors"
    Sheets("DataSheet").Cells(je, 1) = "end of endogenous"
    Sheets("DataSheet").Cells(js, 1) = "end of lines"
Case Is = "sort":
    For i = 2 To je
        Sheets("DataSheet").Cells(i, mycol) = i - 1
    Next
    Range(Cells(2, mycol), Cells(js + 1, mycol)).HorizontalAlignment = xlCenter
Case Is = "Receipts":
Set myrange = Range(Cells(2, mycol), Cells(js, mycol))
    myrange.Formula = "='SAM>>'!RC[-2]"
Case Is = "GrossOutput":
    For i = 2 To jc
        Cells(i, mycol).Formula = "='SAM>>'!R" & js + 2 & "C" & i
    Next
Case Is = "pgross":
    Sheets("DataSheet").Range(Cells(2, mycol), Cells(jc, mycol)).FormulaR1C1 = "=RC[-1]/R" & js + 1 & "C[-1]"
Case Is = "BaseOutput":
    For i = 2 To je
        Cells(i, mycol).Formula = "=OutImp!R" & jc + 2 & "C" & i
    Next
Case Is = "pbase":
    Sheets("DataSheet").Range(Cells(2, mycol), Cells(je, mycol)).FormulaR1C1 = "=RC[-1]/R" & js + 1 & "C[-1]"
Case Is = "multOutput":
    For i = 2 To je
        Cells(i, mycol).Formula = "='I-S inv'!R" & je + 2 & "C" & i
    Next
    Cells(js + 1, mycol).ClearContents
Case Is = "multBuss":
    For i = 2 To jc
        Cells(i, mycol).Formula = "='I-S inv'!R" & je + 3 & "C" & i
    Next
    Cells(js + 1, mycol).ClearContents

Case Is = "directOutput":
    mycol2 = Application.VLookup("Exogenous", ActiveWorkbook.Names("Structure").RefersToRange, 2, False)
    Range(Cells(2, mycol), Cells(jc, mycol)).FormulaR1C1 = "=RC[" & Val(mycol2 - mycol) & "]"
Case Is = "indirectOutput":
    mycol2 = Application.VLookup("BaseOutput", ActiveWorkbook.Names("Structure").RefersToRange, 2, False)
    mycol3 = Application.VLookup("directOutput", ActiveWorkbook.Names("Structure").RefersToRange, 2, False)
    Range(Cells(2, mycol), Cells(je, mycol)).FormulaR1C1 = "=if(RC[" & Val(mycol2 - mycol) & "]>0,RC[" & Val(mycol2 - mycol) & "]-RC[" & Val(mycol3 - mycol) & "],0)"
Case Is = "grossEmpl":
    For i = 2 To jc
        Cells(i, mycol).Formula = "=inputEMPL!R" & i & "C3"
    Next
    Set myrange = Range(Cells(2, mycol), Cells(js + 1, mycol))
    addNameRange "employment", myrange
    Sheets("inputEMPL").Visible = False
Case Is = "pgrossEmpl":
    Sheets("DataSheet").Range(Cells(2, mycol), Cells(jc, mycol)).FormulaR1C1 = "=RC[-1]/R" & js + 1 & "C[-1]"
Case Is = "BaseEmpl":
    For i = 2 To je
        Cells(i, mycol).Formula = "=EmpImp!R" & jc + 2 & "C" & i
    Next
Case Is = "pBaseEmpl":
    Sheets("DataSheet").Range(Cells(2, mycol), Cells(je, mycol)).FormulaR1C1 = "=RC[-1]/R" & js + 1 & "C[-1]"
Case Is = "jobsOutput":
    For i = 2 To jc
        Cells(i, mycol).Formula = "='EmpMult'!R" & jc + 2 & "C" & i
    Next
    Cells(js + 1, mycol).ClearContents
Case Is = "multEmpl":
    mycol2 = Application.VLookup("jobsOutput", ActiveWorkbook.Names("Structure").RefersToRange, 2, False)
    mycol3 = Application.VLookup("jobsCoefficient", ActiveWorkbook.Names("Structure").RefersToRange, 2, False)
    Range(Cells(2, mycol), Cells(jc, mycol)).FormulaR1C1 = _
        "=if(RC[" & Val(mycol3 - mycol) & "]>0,RC[" & Val(mycol2 - mycol) & "]/RC[" & Val(mycol3 - mycol) & "]," & Chr(34) & Chr(34) & ")"
    Cells(js + 1, mycol).ClearContents
Case Is = "directEmpl":
    Range(Cells(2, mycol), Cells(jc, mycol)).FormulaR1C1 = "=employment/linetotals * Exogenous"
Case Is = "indirectEmpl":
    mycol2 = Application.VLookup("BaseEmpl", ActiveWorkbook.Names("Structure").RefersToRange, 2, False)
    mycol3 = Application.VLookup("directEmpl", ActiveWorkbook.Names("Structure").RefersToRange, 2, False)
    Range(Cells(2, mycol), Cells(je, mycol)).FormulaR1C1 = "=if(RC[" & Val(mycol2 - mycol) & "]>0,RC[" & Val(mycol2 - mycol) & "]-RC[" & Val(mycol3 - mycol) & "],0)"
Case Is = "JobsCoefficient":
    Range(Cells(2, mycol), Cells(jc, mycol)).FormulaR1C1 = "=10^3*" & dollars & "*employment/linetotals"
    Cells(js + 1, mycol).ClearContents
Case Is = "GrossWages":
    For i = 2 To jc
        Cells(i, mycol).Formula = "='SAM>>'!R" & jw & "C" & i & ""
    Next
    Set myrange = Range(Cells(2, mycol), Cells(js + 1, mycol))
    addNameRange "wages", myrange
    Range(Cells(2, mycol), Cells(js + 1, mycol)).NumberFormat = "_($* #,##0_);_($* (#,##0);_($* ""-""??_);_(@_)"
Case Is = "pgrossWages":
    Sheets("DataSheet").Range(Cells(2, mycol), Cells(jc, mycol)).FormulaR1C1 = "=RC[-1]/R" & js + 1 & "C[-1]"
Case Is = "BaseWages":
    For i = 2 To je
        Cells(i, mycol).Formula = "=WAgeImp!R" & jc + 2 & "C" & i
    Next
Case Is = "multWages":
    For i = 2 To je
        Cells(i, mycol).Formula = "=WAgeMult!R" & jc + 2 & "C" & i
    Next
Case Is = "wagesperwage":
    mycol2 = Application.VLookup("multWages", ActiveWorkbook.Names("Structure").RefersToRange, 2, False)
    Range(Cells(2, mycol), Cells(jc, mycol)).FormulaR1C1 = _
        "=IF(WageCoefficient>0,RC[" & Val(mycol2 - mycol) & "]/WageCoefficient," & Chr(34) & Chr(34) & ")"
Case Is = "pbaseWages":
    Sheets("DataSheet").Range(Cells(2, mycol), Cells(je, mycol)).FormulaR1C1 = "=RC[-1]/R" & js + 1 & "C[-1]"
Case Is = "directWages":
'= WageCoefficient * exegonous
'lets do the coefficient (i.e. wages/output)
    mycol2 = Application.VLookup("WageCoefficient", ActiveWorkbook.Names("Structure").RefersToRange, 2, False)
    Range(Cells(2, mycol2), Cells(jc, mycol2)).FormulaR1C1 = "=Wages/linetotals"
    addNameRange "WageCoefficient", Range(Cells(2, mycol2), Cells(jc, mycol2))
    Cells(js + 1, mycol2).ClearContents
'now the direct empl
    Range(Cells(2, mycol), Cells(jc, mycol)).FormulaR1C1 = "=WageCoefficient * Exogenous"
Case Is = "indirectWages":
    mycol2 = Application.VLookup("BaseWages", ActiveWorkbook.Names("Structure").RefersToRange, 2, False)
    mycol3 = Application.VLookup("directWages", ActiveWorkbook.Names("Structure").RefersToRange, 2, False)
    Range(Cells(2, mycol), Cells(je, mycol)).FormulaR1C1 = "=if(RC[" & Val(mycol2 - mycol) & "]>0,RC[" & Val(mycol2 - mycol) & "]-RC[" & Val(mycol3 - mycol) & "],0)"
Case Is = "WageOutput":
    Range(Cells(2, mycol), Cells(jc, mycol)).FormulaR1C1 = "=wages/linetotals"
    Cells(js + 1, mycol).ClearContents
Case Is = "WageCoefficient":
    'already done
Case Is = "Exogenous":
    For i = 2 To js
        Cells(i, mycol).Formula = "=sum('SAM>>'!R" & i & "C" & je + 1 & ":R" & i & "C" & js & ")"
    Next
    Range(Cells(2, mycol), Cells(js + 1, mycol)).NumberFormat = "_($* #,##0_);_($* (#,##0);_($* ""-""??_);_(@_)"
    addNameRange id, "=DataSheet!R2C" & mycol & ":R" & js & "C" & mycol
Case Is = "Endogenous": myFormula = ""
    For i = 2 To js
        Cells(i, mycol).Formula = "=sum('SAM>>'!R" & i & "C2:R" & i & "C" & je & ")"
    Next
    Range(Cells(2, mycol), Cells(js + 1, mycol)).NumberFormat = "_($* #,##0_);_($* (#,##0);_($* ""-""??_);_(@_)"
    addNameRange id, "=DataSheet!R2C" & mycol & ":R" & js & "C" & mycol
Case Is = "GrossVA"
    For i = 2 To jc
        Cells(i, mycol).Formula = "=sum('SAM>>'!R" & jw & "C" & i & ":R" & jw + 3 & "C" & i & ")"
    Next
    addNameRange id, "=DataSheet!R2C" & mycol & ":R" & jc & "C" & mycol
Case Is = "pVA":
    Sheets("DataSheet").Range(Cells(2, mycol), Cells(jc, mycol)).FormulaR1C1 = "=RC[-1]/R" & js + 1 & "C[-1]"
Case Is = "BaseVA":
    For i = 2 To je
        Cells(i, mycol).Formula = "=VAImp!R" & jc + 2 & "C" & i
    Next
Case Is = "multVA":
    For i = 2 To je
        Cells(i, mycol).Formula = "=VAMult!R" & jc + 2 & "C" & i
    Next
Case Is = "pbaseVA":
    Sheets("DataSheet").Range(Cells(2, mycol), Cells(je, mycol)).FormulaR1C1 = "=RC[-1]/R" & js + 1 & "C[-1]"
Case Is = "directVA":
'= VACoefficient * exegonous
'lets do the coefficient (i.e. VA/output)
    mycol2 = Application.VLookup("VACoefficient", ActiveWorkbook.Names("Structure").RefersToRange, 2, False)
    Range(Cells(2, mycol2), Cells(jc, mycol2)).FormulaR1C1 = "=GrossVA/linetotals"
    addNameRange "VACoefficient", Range(Cells(2, mycol2), Cells(jc, mycol2))
    Cells(js + 1, mycol2).ClearContents
'now the direct VA
    Range(Cells(2, mycol), Cells(jc, mycol)).FormulaR1C1 = "=VACoefficient * Exogenous"
Case Is = "indirectVA":
    mycol2 = Application.VLookup("BaseVA", ActiveWorkbook.Names("Structure").RefersToRange, 2, False)
    mycol3 = Application.VLookup("directVA", ActiveWorkbook.Names("Structure").RefersToRange, 2, False)
    Range(Cells(2, mycol), Cells(je, mycol)).FormulaR1C1 = "=if(RC[" & Val(mycol2 - mycol) & "]>0,RC[" & Val(mycol2 - mycol) & "]-RC[" & Val(mycol3 - mycol) & "],0)"
Case Is = "VAOutp":
    mycol2 = Application.VLookup("GrossVA", ActiveWorkbook.Names("Structure").RefersToRange, 2, False)
    Range(Cells(2, mycol), Cells(jc, mycol)).FormulaR1C1 = "=RC" & mycol2 & "/linetotals"
    Cells(js + 1, mycol).Clear
Case Is = "VACoefficient":
    'alrady done
Case Is = "EndogenousPurchases":
    For i = 2 To je
        Cells(i, mycol).Formula = "='S_matrix'!R" & je + 2 & "C" & i
    Next
    Range(Cells(2, mycol), Cells(js + 1, mycol)).HorizontalAlignment = xlCenter
    Cells(js + 1, mycol).ClearContents
End Select
End Sub
Sub datasheetformat()
    Sheets("datasheet").Select
    For i = 1 To 3
        Select Case i
        Case Is = 1:
            Set myrange = Range(Cells(2, 1), Cells(jc, 1))
            datasheetFormat_2 ("sectors")
            Set myrange = Range(Cells(1, 1), Cells(jc, Range("1:1").Find("").Column - 1))
            myrange.Interior.Pattern = xlSolid
            myrange.Interior.TintAndShade = 0
        Case Is = 2: Set myrange = Range(Cells(jc + 1, 1), Cells(je, 1))
            datasheetFormat_2 ("endogenous")
            Set myrange = Range(Cells(jc + 1, 1), Cells(je, Range("1:1").Find("").Column - 1))
            myrange.Interior.Pattern = xlSolid
            myrange.Interior.TintAndShade = -4.99893185216834E-02
        Case Is = 3:  Set myrange = Range(Cells(je + 1, 1), Cells(js, 1))
            datasheetFormat_2 ("exogenous")
            Set myrange = Range(Cells(je + 1, 1), Cells(js, Range("1:1").Find("").Column - 1))
            myrange.Interior.Pattern = xlSolid
            myrange.Interior.TintAndShade = -0.149998474074526
        End Select
    Next
    Range("A:A").EntireColumn.AutoFit
    boxin Range(Cells(1, 1), Cells(1, Range("1:1").Find("").Column - 1)), True
    Set myrange = Range(Cells(js + 1, 2), Cells(js + 1, Range("1:1").Find("").Column - 1))
        myrange.Borders(xlEdgeTop).LineStyle = xlContinuous
        myrange.Borders(xlEdgeTop).Weight = xlThin
        myrange.Borders(xlEdgeBottom).LineStyle = xlContinuous
        myrange.Borders(xlEdgeBottom).Weight = xlThin
        Calculate
        myrange.Font.Size = 8
    Range("D2").Select
    ActiveWindow.freezepanes = True
End Sub
Sub datasheetFormat_2(myval)
        With myrange
            .MergeCells = True
            .Value = myval
            .Font.Size = 12
            .Font.Bold = True
            .Orientation = -90
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
        End With
End Sub
