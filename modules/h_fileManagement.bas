Attribute VB_Name = "h_fileManagement"
Option Explicit
Public Sub FAQ()
    Sheets("FAQ").Select
    Cells(1, 1).Select
End Sub

Public Sub saveCopy(how)
    Dim fname, thisfile, response, shortname As Variant
    Select Case how
    Case Is = 1:
'--------------------------------------------
'save a functional copy (i.e. with macros)
        fname = Application.GetSaveAsFilename(InitialFileName:=ThisWorkbook.Path & "\SAM_" & Format(Date, "yyyymmdd") & ".xlsm", _
            fileFilter:=" Excel Macro Enabled Workbook (*.xlsm), *.xlsm", _
            FilterIndex:=1, title:="File as Macro-enabled to preserve functionaility...")
        If fname = False Then Exit Sub
        If CStr(fname) = ActiveWorkbook.FullName Then
        'apparently to just updating the current file but using save copy...
            ActiveWorkbook.Save
            Exit Sub
        End If
    'set the macro warning as first sheet and saveCopy
        Application.ScreenUpdating = False
        Sheets("macrohelp").Visible = True
            Sheets("macrohelp").Select
            Range("A1").Select
        ActiveWorkbook.SaveCopyAs filename:=CStr(fname)
    'reset visibility to false for the current file
        Sheets("macrohelp").Visible = False
            Sheets("tools").Select
            Range("A1").Select
        Application.ScreenUpdating = True
    Case Is = 2:
'--------------------------------------------
'save an archive copy without macros
        thisfile = ActiveWorkbook.FullName
        If ActiveWorkbook.Saved = False Then
            response = MsgBox("Save your last changes?" & vbCrLf & _
                "You will loose the changes if you continue...", vbYesNoCancel)
                If response = vbCancel Then Exit Sub
                If response = 6 Then ActiveWorkbook.Save
        End If
        fname = Application.GetSaveAsFilename(InitialFileName:=ThisWorkbook.Path & "\SAMarchive_" & Format(Date, "yyyymmdd") & ".xlsx", _
            fileFilter:=" Excel Workbook (*.xlsx), *.xlsx", _
            FilterIndex:=1, title:="Copy as spreadsheet only...")
        If fname = False Then Exit Sub
    'prepare for save and strip the file of functionality
        Sheets("tools").Select
        Cells(1, 1).Select
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
        prepareFunctionalCopy
        ActiveWorkbook.SaveAs filename:=fname, FileFormat:=51
        shortname = ActiveWorkbook.Name 'reset fname to short name of archive copy
        MsgBox "Your archive copy is saved as:" & vbCr & shortname
        Shell "C:\WINDOWS\explorer.exe /select," & fname & """", vbNormalFocus
        Application.Quit
    End Select
End Sub

Sub prepareFunctionalCopy()
'strip sheets, buttons and other macro functionality
    Sheets("MacroHelp").delete
    Sheets("structure").delete
'cleanup the tools sheet
    Sheets("tools").Select
        ActiveSheet.Unprotect
        Range("A2:F27").Select
        Selection.ClearContents
        Range("A2:F27").Select
        Selection.ClearContents
        Dim shp As Shape
            For Each shp In ActiveSheet.Shapes
                Debug.Print shp.Type
                If Not shp.OnAction = "" Or shp.Type = 8 Then
                    shp.Select
                    shp.delete
                End If
            Next
        Range("D5").Value = "This is an archive copy without functionality"
        ActiveSheet.Protect
'delete visual basic modules
    Dim mymodules As New Cmodule
    mymodules.delete (1) 'regular bas modules
    mymodules.delete (2) 'regular bas modules
    mymodules.delete (3) ' forms
    Set mymodules = Nothing
End Sub

Public Function FileFolderExists(strFullPath As String) As Boolean
    On Error GoTo myerror
    If Not dir(strFullPath, vbDirectory) = vbNullString Then FileFolderExists = True
    Exit Function
myerror:
    FileFolderExists = False
End Function
'============================================
'currently not being used
Public Sub SelectSaveFileName()
Dim fname, fileformatvalue As Variant
    If (ActiveWorkbook.Saved = False) Then
        If MsgBox("Save your last changes?", vbYesNo) = 6 Then ActiveWorkbook.Save
    End If
    fname = Application.GetSaveAsFilename(InitialFileName:="SAMcopy_" & Format(Date, "yyyymmdd") & ".xlsm", _
    fileFilter:= _
        " Excel Macro Enabled Workbook (*.xlsm), *.xlsm," & _
        " Excel Macro Free Workbook (*.xlsx), *.xlsx," & _
        " Excel 2000-2003 Workbook (*.xls), *.xls," & _
        " Excel Macro Enabled Template (*.xltm), *.xltm", _
        FilterIndex:=1, title:="File as Macro-enabled to preserve functionaility...")
    If fname = False Then Exit Sub
    Select Case LCase(Right(fname, Len(fname) - InStrRev(fname, ".", , 1)))
        Case "xls": fileformatvalue = 56
        Case "xlsx": fileformatvalue = 51
        Case "xlsm": fileformatvalue = 52
        Case "xltm": fileformatvalue = 53
            fname = "SAMtemplate" & Format(Date, "yymmdd") & ".xltm"
            Application.ScreenUpdating = False
            Sheets("MacroHelp").Visible = True
            Sheets("MacroHelp").Select
            Range(Cells(2, 2), Cells(2, 3)).Select
            Application.ScreenUpdating = True
        Case Else: fileformatvalue = 0
    End Select
    ActiveWorkbook.SaveAs filename:=CStr(fname), FileFormat:=fileformatvalue
End Sub



