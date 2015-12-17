Attribute VB_Name = "support_functions"
Public Function hilite(thisrange, how)
'lastModified 20100518
    If how > 0 Then
        thisrange.Interior.Color = IIf(how = 2, 9486586, 65535)
        thisrange.Interior.Pattern = xlSolid
    Else
        thisrange.Interior.Pattern = xlNone
    End If
End Function

'nameranges=============
Public Function addNameRange(myname, namerange)
'lastModified 20100518
    deleteNameRange (myname)
    ActiveWorkbook.Names.add Name:=myname, RefersToR1C1:=namerange
End Function
Public Function deleteNameRange(myname)
'lastModified 20100518
    On Error Resume Next
    ActiveWorkbook.Names(myname).delete
End Function
'sheet management==========
Public Function replaceSheet(what)
'lastModified 20100518
    On Error Resume Next
    Application.DisplayAlerts = False
    Sheets(what).delete
    Application.DisplayAlerts = True
    Sheets.add After:=Sheets(Sheets.Count)
    Sheets(Sheets.Count).Select
    Sheets(Sheets.Count).Name = what
End Function
Public Function checkname(myname)
    On Error Resume Next
    checkname = (ActiveWorkbook.Names(myname) <> "")
    On Error GoTo 0
End Function

Public Function golastcell()
    Selection.SpecialCells(xlCellTypeLastCell).Select
    mycol = ActiveCell.Column - 1
    ActiveCell.Offset(0, -mycol).Range("A1").Select
End Function

Public Function WorksheetExists(ByVal WorksheetName As String) As Boolean
'lastModified 20100518
    On Error Resume Next
    WorksheetExists = (Sheets(WorksheetName).Name <> "")
End Function

Public Function reset()
    Application.ScreenUpdating = True
End Function

Public Function tellshade()
    InputBox "", , Selection.Interior.Color
End Function

Public Function showinfo(mytext, Optional CloseAfter As Boolean, Optional mypause As Integer, Optional doYield As Boolean)
    If IsMissing(doYield) Then doYield = True
    If IsMissing(CloseAfter) Then CloseAfter = False
    frm_info.progress.Caption = Trim(mytext)
    frm_info.show
    frm_info.Repaint
    Timer mypause, doYield
    If CloseAfter = True Then
        frm_info.progress = ""
        frm_info.Hide
    End If
End Function
Public Function Timer(PauseTime, Optional mystart, Optional doYield)
If IsMissing(doYield) Then doYield = True
Dim finish, TotalTime
    start = Format(Now(), "nnss")
    Do While Val(Format(Now(), "nnss")) < start + PauseTime
        If doYield = True Then DoEvents   ' Yield to other processes.
    Loop
End Function

Public Function goTools()
    Sheets("tools").Select
End Function
