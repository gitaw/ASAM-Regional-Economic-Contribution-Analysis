Attribute VB_Name = "i_matrices"
Sub ClearMatrices(Optional start)
'lastModified 20100518
    Application.ScreenUpdating = False
    On Error GoTo myerror
    If IsMissing(start) Then
        If MsgBox("This will delete all data from this workbook, preparing it for a new SAM analysis; are you sure?", vbCritical + vbOKCancel, "Alert") = vbCancel Then Exit Sub
        start = 0
    End If
    Application.DisplayAlerts = False
    For i = start To 19
        sheetname = mysheets(i)
            If WorksheetExists(sheetname) Then
                showinfo "delete [" & sheetname & "]"
                Sheets(sheetname).delete
            End If
    Next
    showinfo "Finishing..."
    If Not WorksheetExists("inputEMPL") Then
        Sheets.add After:=Sheets(Sheets.Count)
        Sheets(Sheets.Count).Name = "inputEMPL"
        Cells(1, 1) = "type"
        Cells(1, 2) = "Institutions"
        Cells(2, 2) = "(optional)"
        Cells(1, 3) = "Gross Employment"
        hilite Cells(2, 3), 1
        Range("A:C").EntireColumn.AutoFit
        Cells(4, 4) = "1. To manually paste your employment data from the Access database, start pasting at cell C1"
        
        Cells(1, 4) = " <<<Label (optional)."
        Cells(2, 4) = " <<<Numerical data start at c2 and down...."
        Cells(3, 4) = "The descriptions (optional; they should be the same as in the sam) should be in column A"
        Cells(5, 4) = "2. Then return to the tools-sheet and click [Create Matrices]"
        Rows(1).Borders(xlBottom).LineStyle = xlContinuous
        Rows(1).Borders(xlBottom).Weight = xlThin
        boxin Columns("A:B"), 1
        Columns(2).Borders(xlEdgeLeft).LineStyle = xlContinuous
        Columns(2).Borders(xlEdgeLeft).Weight = xlThin
        Columns(3).Borders(xlEdgeLeft).LineStyle = xlContinuous
        Columns(3).Borders(xlEdgeLeft).Weight = xlThin
        Columns(3).Borders(xlEdgeRight).LineStyle = xlContinuous
        Columns(3).Borders(xlEdgeRight).Weight = xlThin
        formatManualEntry
        Sheets("inputEMPL").Visible = False
    End If
    If Not WorksheetExists("SAM>>") Then
        Sheets.add After:=Sheets(Sheets.Count)
        Sheets(Sheets.Count).Name = "SAM>>"
        Columns(1).Borders(xlEdgeRight).LineStyle = xlContinuous
        Columns(1).Borders(xlEdgeRight).Weight = xlThin
        hilite Cells(1, 1), 1
        Rows(1).Borders(xlBottom).LineStyle = xlContinuous
        Rows(1).Borders(xlBottom).Weight = xlThin
        Cells(5, 3) = "1. To manually paste your SAM data from the access database, start pasting from cell A1"
        boxin Cells(2, 2), False
        Cells(2, 3) = "<< Numerical data start here (B2 and across)...."
        Cells(1, 2) = "<< Sector labels should be in column A"
        Cells(6, 3) = "2. Select the sheet [inputEmpl] to input your employment data"
        Cells(7, 3) = "3. Then return to the tools-sheet and click [Create Matrices]"
        formatManualEntry
    End If
myend:
    Sheets("tools").Select
    Application.ScreenUpdating = True
    Application.Calculation = xlAutomatic
    Application.DisplayAlerts = True
    showinfo "", True
    Exit Sub
myerror:
    GoTo myend
End Sub

Function initSAM()
    initSAM = True
'lastModified 20100725
    Sheets("SAM>>").Select
    Cells(1, 1) = "Receipts/payments in: " & receipts
    ActiveSheet.Range(Cells(2, 1), Cells(js, 1)).Copy
    Cells(1, 2).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=True
    Cells(js + 2, 1).FormulaR1C1 = "Sum"
    Set myrange = Range(Cells(js + 2, 2), Cells(js + 2, js))
        myrange.FormulaR1C1 = "=SUM('SAM>>'!R2C:R" & js & "C)"
        boxin myrange, True
    Cells(1, js + 2) = "Sum"
    Set myrange = Range(Cells(2, js + 2), Cells(js, js + 2))
        myrange.FormulaR1C1 = "=SUM('SAM>>'!RC2:RC" & js & ")"
        boxin myrange, True
        addNameRange "linetotals", myrange
        myrange.Copy
    Cells(js + 3, 1) = "Transposed"
    Cells(js + 3, 2).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=True
    Cells(js + 4, 1) = "Sum-check"
    Range(Cells(js + 4, 2), Cells(js + 4, js + 1)).FormulaR1C1 = "=r[-1]c - r[-2]C"
    Cells(js + 2, js + 2).FormulaR1C1 = "=sum(R2C" & js + 2 & ":R" & js & "C" & js + 2 & ")"
'do some formatting
    Sheets("sam>>").Cells(2, 2).Select
    samFormat
    If sumcheck = True Then initSAM = False
'fill 1s for the TY matrix
    Range(Cells(js + 1, 2), Cells(js + 1, je)).Value = 1
End Function

Sub createMatrices()
'create matrices
    For i = 2 To 15
        domatrix (mysheets(i))
        mystop = IIf(i < 9, je, jc)
        sectorLabels mystop
        generalFormat
    Next
End Sub

Function domatrix(what)
'lastModified 20151122: corrected EmpMult multiplication, which was wrongly multiplied by the dollar variable
'                       making it a factor [dollar] off if dollar>0

    replaceSheet (what)
    Set myrange = Range(Cells(2, 2), Cells(je, je))
    If what = "VAmult" Then Stop
    Select Case what
        Case Is = "I_matrix":
            For i = 2 To je
                Cells(i, i) = "1"
            Next
        Case Is = "S_matrix":
            myrange = "='SAM>>'!RC/SUM('SAM>>'!R2C:R" & js & "C)"
            Cells(je + 2, 1) = "Local Purchases"
            Range(Cells(je + 2, 2), Cells(je + 2, je)) = "=SUM(R[" & -je & "]C:R[-2]C)"
            Cells(je + 3, 1) = "Exogenous purchases"
            Range(Cells(je + 3, 2), Cells(je + 3, je)) = "=1-R[-1]C"
            Range(Cells(2, 2), Cells(je + 3, je)).NumberFormat = "0.00%"
            Exit Function
        Case Is = "I-S":
            myrange = "=I_matrix!RC-S_matrix!RC"
            myrange.NumberFormat = "0.0000"
        Case Is = "I-S inv":
            myrange.FormulaArray = "=MINVERSE('I-S'!RC:R[" & je - 2 & "]C[" & je - 2 & "])"
            addNameRange "inverse", "='I-S inv'!R2C2:R" & je & "C" & je
            Cells(je + 2, 1) = "Total output multiplier"
            Range(Cells(je + 2, 2), Cells(je + 2, je)) = "=SUM(R[" & -je & "]C:R[-2]C)"
            Cells(je + 3, 1) = "Business multiplier "
            Range(Cells(je + 3, 2), Cells(je + 3, jc)).FormulaR1C1 = "=SUM(R[" & -je - 1 & "]C:R[" & -je + jc - 3 & "]C)"
            ActiveSheet.Range(Cells(2, 2), Cells(je + 3, je)).NumberFormat = "0.0000"
        Case Is = "TY(int)":
            myrange.FormulaArray = "=MMULT(Exogenous,'SAM>>'!R[" & js - 1 & "]C2:R[" & js - 1 & "]C[" & je - 2 & "])"
            myrange.NumberFormat = "_($* #,##0_);_($* (#,##0);_($* ""-""??_);_(@_)"
        Case Is = "TY":
            addNameRange "TY", "='TY'!R2C2:R" & je & "C" & je
            myrange.FormulaR1C1 = "=I_matrix!RC*'TY(int)'!RC"
            myrange.NumberFormat = "0.0000"
        Case Is = "Z":
            myrange.FormulaArray = "=MMULT(inverse,TY)"
            addNameRange "Z", "='Z'!R2C2:R" & je & "C" & je
            myrange.NumberFormat = "_($* #,##0_);_($* (#,##0);_($* ""-""??_);_(@_)"
        Case Is = "OutImp":
            Range(Cells(2, 2), Cells(jc, je)) = "=Z!B2*1"
            Cells(jc + 2, 1) = "total"
            ActiveSheet.Range(Cells(jc + 2, 2), Cells(jc + 2, je)) = "=sum(R2C:R" & jc & "C)"
            Cells(jc + 3, 1) = "grand total"
            Cells(jc + 3, 2) = "=sum(R" & jc + 2 & "C:R" & jc + 2 & "C" & je & ")"
            Range(Cells(2, 2), Cells(jc + 3, je)).NumberFormat = "_($* #,##0_);_($* (#,##0);_($* ""-""??_);_(@_)"
        Case Is = "WageImp":
            Range(Cells(2, 2), Cells(jc, je)).FormulaR1C1 = "=Z!RC * wages/linetotals"
            Cells(jc + 2, 1) = "total"
            Range(Cells(jc + 2, 2), Cells(jc + 2, je)) = "=sum(R2C:R" & jc & "C)"
            Cells(jc + 3, 1) = "sector total"
            Cells(jc + 3, 2) = "=sum(R" & jc + 2 & "C:R" & jc + 2 & "C" & je & ")"
            Range(Cells(2, 2), Cells(jc + 3, je)).NumberFormat = "_($* #,##0_);_($* (#,##0);_($* ""-""??_);_(@_)"
        Case Is = "EmpImp":
            Range(Cells(2, 2), Cells(jc, je)) = "=Z!RC*(employment/'SAM>>'!RC" & js + 2 & ")"
            Cells(jc + 2, 1) = "Sector Total"
            ActiveSheet.Range(Cells(jc + 2, 2), Cells(jc + 2, je)) = "=sum(R2C:R" & jc & "C)"
            ActiveSheet.Cells(jc + 3, 1) = "Grand Total"
            Cells(jc + 3, 2) = "=sum(R" & jc + 2 & "C:R" & jc + 2 & "C" & je & ")"
            Range(Cells(2, 2), Cells(jc + 3, je)).NumberFormat = "_(* #,##0_);_(* (#,##0);_(* ""-""??_);_(@_)"
        Case Is = "VAImp":
            Range(Cells(2, 2), Cells(jc, je)) = "=Z!RC*(GrossVA/'SAM>>'!RC" & js + 2 & ")"
            Cells(jc + 2, 1) = "Sector Total"
            ActiveSheet.Range(Cells(jc + 2, 2), Cells(jc + 2, je)) = "=sum(R2C:R" & jc & "C)"
            ActiveSheet.Cells(jc + 3, 1) = "Grand Total"
            Cells(jc + 3, 2) = "=sum(R" & jc + 2 & "C:R" & jc + 2 & "C" & je & ")"
            Range(Cells(2, 2), Cells(jc + 3, je)).NumberFormat = "_(* #,##0_);_(* (#,##0);_(* ""-""??_);_(@_)"
        Case Is = "WageMult":
            Range(Cells(2, 2), Cells(jc, je)) = "='I-S inv'!RC*wages/linetotals"
            Cells(jc + 2, 1) = "wages multiplier"
            ActiveSheet.Range(Cells(jc + 2, 2), Cells(jc + 2, je)) = "=sum(R2C:R" & jc & "C)"
            Range(Cells(2, 2), Cells(jc + 3, je)).NumberFormat = "0.00"
        Case Is = "EmpMult":
            Range(Cells(2, 2), Cells(jc, je)) = "='I-S inv'!RC*1000 *employment/linetotals"
            Cells(jc + 2, 1) = "employment multiplier"
            ActiveSheet.Range(Cells(jc + 2, 2), Cells(jc + 2, je)) = "=sum(R2C:R" & jc & "C)"
            Range(Cells(2, 2), Cells(jc + 3, je)).NumberFormat = "0.00"
        Case Is = "VAMult":
            Range(Cells(2, 2), Cells(jc, je)) = "='I-S inv'!RC *GrossVA/linetotals"
            Cells(jc + 2, 1) = "Value Added multiplier"
            ActiveSheet.Range(Cells(jc + 2, 2), Cells(jc + 2, je)) = "=sum(R2C:R" & jc & "C)"
            Range(Cells(2, 2), Cells(jc + 3, je)).NumberFormat = "0.00"
        End Select
End Function

   
Function sumcheck()
'lastModified 20100725
    Sheets("SAM>>").Select
    Max = 100
    Cells(js + 5, 1).FormulaR1C1 = "=AVERAGE(r[-2]c[1]:r[-2]c[" & js - 1 & "])"
    myavg = Cells(js + 5, 1).Value
    If Max < 0.00001 * myavg Then Max = Round(0.00001 * myavg, 0)
    sumcheck = False
    Cells(js + 5, 1).ClearContents
    For i = 2 To js + 1
        If Abs(Cells(js + 4, i).Value) > Max Then
            If MsgBox("The sumcheck for " & Cells(1, i).Value & _
                " is greater than " & Max & "; do you want to continue?" & vbCrLf & vbCrLf & _
                "If you click [YES], this  will be ignored and the SAM construction will continue...", vbCritical + vbYesNo) = vbNo Then
                    Cells(js + 4, i).Select
                sumcheck = True
                Exit Function
            End If
        End If
    Next
End Function

Function sectorLabels(mystop)
'lastModified 20100725
    Range(Cells(1, 1), Cells(mystop, 1)).Formula = "='SAM>>'!rc"
    Range(Cells(1, 1), Cells(1, je)).Formula = "='SAM>>'!rc"
    Cells(2, 2).Select
    ActiveWindow.freezepanes = True
End Function


Sub samFormat()
'lastModified 20151107
    Sheets("SAM>>").Select
    boxin Range(Cells(1, 1), Cells(1, js)), 2
    boxin Range(Cells(1, 1), Cells(js + 4, 1)), 0
    boxin Range(Cells(1, 1), Cells(js, 1)), 2
    boxin Range(Cells(js + 2, 1), Cells(js + 2, js + 2)), 1
    boxin Range(Cells(js + 3, 1), Cells(js + 3, js)), 1
    boxin Range(Cells(js + 4, 1), Cells(js + 4, js)), 1
    boxin Range(Cells(1, js + 2), Cells(js + 4, js + 2)), 1
    generalFormat
    Set myrange = Rows(1)
        Rows(1).VerticalAlignment = xlCenter
        Rows(1).WrapText = True
        Rows(1).HorizontalAlignment = xlCenter
    Set myrange = Columns(1)
    Range(Cells(2, 2), Cells(js + 4, js + 2)).NumberFormat = "_($* #,##0_);_($* (#,##0);_($* ""-""??_);_(@_)"
    Set myrange = Range(Cells(js + 1, 1), Cells(js + 1, js + 2))
        hilite myrange, (2)
        myrange.RowHeight = 2
    Set myrange = Range(Cells(1, js + 1), Cells(js + 4, js + 1))
        hilite myrange, 2
        myrange.ColumnWidth = 0.2
    Cells(1, 1).VerticalAlignment = xlBottom
    Rows(1).HorizontalAlignment = xlLeft
    addGotoToolsBtn nleft:=10
End Sub

