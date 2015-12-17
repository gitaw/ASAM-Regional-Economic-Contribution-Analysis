Attribute VB_Name = "b_createWorkSheets"
'==============================================
'This module directs the creation of the following worksheets and matrices.
'NOTE: The [mysheets]function itself is actually called only in module i_matrices
'      and placed here to show the overview
'==============================================
Function mysheets(i)
'lastModified 20100518
    Select Case i
        Case Is = 0: mysheets = "SAM>>"
        Case Is = 1: mysheets = "inputEmpl"
        Case Is = 2: mysheets = "I_matrix"
        Case Is = 3: mysheets = "S_matrix"
        Case Is = 4: mysheets = "I-S"
        Case Is = 5: mysheets = "I-S inv"
        Case Is = 6: mysheets = "TY(int)"
        Case Is = 7: mysheets = "TY"
        Case Is = 8: mysheets = "Z"
        Case Is = 9: mysheets = "OutImp"
        Case Is = 10: mysheets = "WageImp"
        Case Is = 11: mysheets = "EmpImp"
        Case Is = 12: mysheets = "VAImp"
        Case Is = 13: mysheets = "WageMult"
        Case Is = 14: mysheets = "EmpMult"
        Case Is = 15: mysheets = "VAMult"
        Case Is = 16: mysheets = "DataSheet"
        Case Is = 17: mysheets = "OutputTable"
        Case Is = 18: mysheets = "Chart(pie)"
        Case Is = 19: mysheets = "Chart(bar)"
    End Select
End Function
'==============================================
'   the following the functions create the
'   datasheets and matrices (in two steps)
'==============================================
Public Sub createSam()
'lastModified 20151107
    ClearMatrices (2) 'clears worksheets from sheets, starting with #2
'simple check: if the template is never updated with real data, cell(2,2)is text
    If (Sheets("SAM>>").Cells(2, 2) = "" Or Not IsNumeric(Sheets("SAM>>").Cells(2, 2))) Then
        Sheets("tools").Select
        MsgBox "You have not yet entered any SAM data yet..." & vbCrLf & _
        "Manually paste your SAM data from the access database using this cell, or use the option to automatically retrieve the data...", vbOKCancel
        Range(Cells(2, 2), Cells(5, 5)).Select
        Exit Sub
    End If
'if there is no Gross Employment in the [inputEMPL] sheet, then note the limitation of the models
    If Not IsNumeric(Sheets("inputEMPL").Cells(2, 3)) Or Sheets("inputEMPL").Cells(2, 3) = "" Then
            Sheets("inputEMPL").Visible = True
            Sheets("inputEMPL").Select
        If MsgBox("You have not yet entered any employment data; that is fine, but it will leave the Output report and EmpImp sheet incomplete" & vbCrLf & _
            "Otherwise, click [Cancel] and paste the employment numbers in B2(down)...", vbOKCancel) = vbCancel Then
            Cells(2, 3).Select
            Exit Sub
        End If
    End If
'initiate questions to determine the sam structure (fsam1-4)
    fsam1
 End Sub
 
Sub createSam_continued(Optional msg)
    If IsMissing(msg) Then msg = ""
'lastModified 20151107
'this continues from CreateSam once the structure is known
    On Error GoTo myerror
'freeze panes while screenupdating=true (if =false it is unpredictable)
    Cells(1, 1).Select
    With ActiveWindow
            .freezepanes = False
            .ScrollRow = 1
            .Split = False
            .SplitColumn = 1
            .SplitRow = 1
            .freezepanes = True
    End With
'make it as fast as possible
    Application.ScreenUpdating = False
    Application.Calculation = xlManual
    Application.Calculation = xlCalculationManual
    msg = msg & "Creating/Updating your SAM; depending on the data and computer this may take a few minutes..." & vbCrLf
'but keep the user updated to avoid confusion on whether it works or not
    showinfo msg
        If initSAM = False Then GoTo myerror
        msg = msg & " - preparing..." & vbCrLf
    showinfo msg
        initDatasheet
        msg = msg & " - creating matrices..." & vbCrLf
    showinfo msg
        createMatrices
        msg = msg & " - creating datasheet..." & vbCrLf
    showinfo msg
        finishDatasheet msg
        msg = msg & " - creating output table..." & vbCrLf
    showinfo msg
        OutputTable
        msg = msg & " - arranging worksheets..." & vbCrLf
    showinfo msg
        OutputTable
        reorderSheets
        msg = msg & " - creating charts..." & vbCrLf
    showinfo msg
        addChartSheet ("pie")
    showinfo msg
        addChartSheet ("bar")
'finished
        Sheets("datasheet").Select
        Cells(1, 1).Select
    showinfo msg & "Done!", True, 1
myexit:
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Exit Sub
myerror:
    showinfo "Aborted...", True, 1
    GoTo myexit
End Sub
'==============================================
'   the following functions interact with the
'   user to determine the structure of the SAM
'==============================================
Public Sub fsam1() 'last line js
'lastModified 20151107
    Application.ScreenUpdating = False
    Sheets("SAM>>").Select
    Range(Cells(1, 1), Selection.SpecialCells(xlCellTypeLastCell)).ClearFormats
    generalFormat
    Application.ScreenUpdating = True
    Range("A2").End(xlDown).Select
    myaddress = ActiveCell.AddressLocal
    hilite ActiveCell, 1
    fsam "Is [" & ActiveCell.Row & "] the last SAM line in the column?" & _
        vbCrLf & "(usually [Imports])" & vbCrLf & _
        vbCrLf & vbCrLf & "If not, click/select the appropriate sector/row", "Next...", "fsam2"
End Sub
Public Sub fsam2()  'last endogenous je
'lastModified 20151107
    hilite ActiveCell, 2
    js = ActiveCell.Row
    If Not ActiveCell.AddressLocal = myaddress Then hilite Range(myaddress), 0
    On Error Resume Next
    Range(Cells(2, 1), Cells(js, 1)).Find(what:="house", After:=Cells(js, 1), SearchDirection:=2).Activate
     If err.Number = 0 Then
        ActiveCell.Select
        msg = "Is [" & ActiveCell.Row & "] the last endogenous sector?" & _
        vbCrLf & "(usually the last [HouseHolds] line)" & vbCrLf & _
        vbCrLf & vbCrLf & "If not, click/select the appropriate sector/row"
        hilite ActiveCell, 1
     Else
        Cells(js - 6, 1).Select
        msg = "Please click the last line of the endogenous sectors..."
     End If
    myaddress = ActiveCell.Address
    fsam msg, "Next...", "fsam3"
End Sub
Function fsam3()    'last sector jc
'lastModified 20151107
    hilite ActiveCell, 2
    je = ActiveCell.Row
    If Not ActiveCell.AddressLocal = myaddress Then hilite Range(myaddress), 0
    On Error Resume Next
    Range(Cells(2, 1), Cells(je - 1, 1)).Find(what:="*employee*", After:=[A2]).Activate
    If err.Number = 0 Then
        msg = "...and finally the last industry or public sector that should be included in the regional analysis?" & _
        vbCrLf & vbCrLf & "If not, click/select the appropriate sector/row"
        ActiveCell.Offset(-1, 0).Select
        hilite ActiveCell, 1
     Else
        msg = "Please click the last of the sectors..."
        Cells(je - 4, 1).Select
     End If
    myaddress = ActiveCell.Address
    fsam msg, "Next...", "fsam4"
End Function
Function fsam4()    'wages and value added
'lastModified 20151107
    hilite ActiveCell, 2
    jc = ActiveCell.Row
    jw = 0
    If Not ActiveCell.AddressLocal = myaddress Then hilite Range(myaddress), 0
    frm_structure.Hide
'find the wages/employee compensation line
    On Error Resume Next
    Range(Cells(2, 1), Cells(je - 1, 1)).Find(what:="Employee", After:=ActiveCell, LookIn:=xlFormulas, _
        LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlPrevious, _
        MatchCase:=False, SearchFormat:=False).Activate
    If err.Number = 0 Then
        jw = ActiveCell.Row '5001
        frm_sam.cmd.SetFocus
    Else
        MsgBox "I could not find the [Employee compensation] line; please enter the appropriate row number..."
        frm_sam.compensation.SetFocus
        frm_sam.compensation.SelStart = 0
        frm_sam.compensation.SelLength = 1
    End If
    ji2 = Range(Cells(jw, 1), Cells(js - 1, 1)).Find(what:="propriet").Row '6001
    ji3 = Range(Cells(jw, 1), Cells(js - 1, 1)).Find(what:="property").Row '7001
    ji4 = Range(Cells(jw, 1), Cells(js - 1, 1)).Find(what:="Business Tax").Row '8001
'allow one final check
    frm_sam.samend.Value = js
    frm_sam.endogenous.Value = je
    frm_sam.sectors.Value = jc
    frm_sam.compensation.Value = jw
    frm_sam.income2.Value = ji2
    frm_sam.income3.Value = ji3
    frm_sam.income4.Value = ji4
    frm_sam.dollars.Value = 1000
    frm_sam.show
End Function
    Function fsam(msg, btn, flw)
    'supports the fsam[i] functions by opening the [frm_sam] form
    'lastModified 20151107
        Application.ScreenUpdating = True
        frm_structure.show
        frm_structure.instruction = msg
        frm_structure.cmd.Caption = btn
        frm_structure.follow = flw
    End Function
Function continuefromfsam()
'lastModified 20100518
    js = Val(frm_sam.samend)
    je = Val(frm_sam.endogenous)
    jc = Val(frm_sam.sectors)
    dollars = Val(frm_sam.dollars.Value)
    If Not Range("dollars") = dollars Then
            Sheets("tools").Unprotect
            Range("dollars") = dollars
            Sheets("tools").Protect DrawingObjects:=True, Contents:=True, Scenarios:=True
    End If
    jw = Val(frm_sam.compensation.Value)
    ji2 = Val(frm_sam.income2.Value)
    ji3 = Val(frm_sam.income3.Value)
    ji4 = Val(frm_sam.income4.Value)
    frm_sam.Hide
    createSam_continued
End Function

Sub reorderSheets()
    Sheets("structure").Move Before:=Sheets(2)
    Sheets("tools").Move Before:=Sheets(3)
    Sheets("inputEMPL").Move Before:=Sheets(4)
    Sheets("DataSheet").Move Before:=Sheets(5)
    Sheets("OutputTable").Move Before:=Sheets(6)
    Sheets("tools").Select
End Sub
'==============================================
'   the following functions interact with the
'   tools page
'==============================================
Public Sub SaveFunctionalCopy()
'lastModified 20100518
    saveCopy 1
End Sub
Public Sub SaveArchiveCopy()
'lastModified 20100518
    saveCopy 2
End Sub
Public Sub EditEmployment()
'lastModified 20100518
    Sheets("inputEMPL").Visible = True
    Sheets("inputEMPL").Select
    Cells(2, 2).Select
End Sub
Public Sub ManualInput()
    Sheets("inputEMPL").Visible = True
    Sheets("SAM>>").Visible = True
    Sheets("SAM>>").Select
    MsgBox "To manually paste your data, make sure that the sector descriptions are in column A and the numerical data start in B2..."
End Sub
Public Sub EditStructure()
    Sheets("structure").Visible = True
    Sheets("structure").Select
    Cells(43, 1).Select
End Sub
Public Sub hideStructure()
    Sheets("structure").Visible = False
    Sheets("tools").Select
    Range("a1:f1").Select
End Sub
