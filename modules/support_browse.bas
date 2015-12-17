Attribute VB_Name = "support_browse"
Option Explicit

Private Const BIF_STATUSTEXT = &H4&
Private Const BIF_RETURNONLYFSDIRS = 1
Private Const BIF_DONTGOBELOWDOMAIN = 2
Private Const MAX_PATH = 260

Private Const WM_USER = &H400
Private Const BFFM_INITIALIZED = 1
Private Const BFFM_SELCHANGED = 2
Private Const BFFM_SETSTATUSTEXT = (WM_USER + 100)
Private Const BFFM_SETSELECTION = (WM_USER + 102)

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Long
Private Declare Function SHBrowseForFolder Lib "Shell32" (lpbi As BrowseInfo) As Long
Private Declare Function SHGetPathFromIDList Lib "Shell32" (ByVal pidList As Long, ByVal lpBuffer As String) As Long
Private Declare Function lstrcat Lib "kernel32" Alias "lstrcatA" (ByVal lpString1 As String, ByVal lpString2 As String) As Long

Private Type BrowseInfo
  hWndOwner      As Long
  pIDLRoot       As Long
  pszDisplayName As Long
  lpszTitle      As Long
  ulFlags        As Long
  lpfnCallback   As Long
  lParam         As Long
  iImage         As Long
End Type

Private m_CurrentDirectory As String   'The current directory
'

'This code snippet was taken from http://forum.hardware.fr/hfr/Programmation/VB-VBA-VBS/vb-dossier-fichier-sujet_36927_1.htm
' Steve's home page did not list it
'=====================================================================================
' Browse for a Folder using SHBrowseForFolder API function with a callback
' function BrowseCallbackProc.
'
' This Extends the functionality that was given in the
' MSDN Knowledge Base article Q179497 "HOWTO: Select a Directory
' Without the Common Dialog Control".
'
' After reading the MSDN knowledge base article Q179378 "HOWTO: Browse for
' Folders from the Current Directory", I was able to figure out how to add
' a callback function that sets the starting directory and displays the
' currently selected path in the "Browse For Folder" dialog.
'
' I used VB 6.0 (SP3) to compile this code.  Should work in VB 5.0.
' However, because it uses the AddressOf operator this code will not
' work with versions below 5.0.
'
' This code works in Window 95a so I assume it will work with later versions.
'
' Stephen Fonnesbeck
' steev@xmission.com
' http://www.xmission.com/~steev
' Feb 20, 2000
'
'=====================================================================================
' Usage:
'=====================================================================================
Public Function BrowseForFolder(title As String, StartDir As String) As String
  'Opens a Treeview control that displays the directories in a computer

  Dim lpIDList As Long
  Dim szTitle As String
  Dim sBuffer As String
  Dim tBrowseInfo As BrowseInfo
  m_CurrentDirectory = StartDir & vbNullChar

  szTitle = title
  With tBrowseInfo
    '.hWndOwner = owner.hWnd
    .lpszTitle = lstrcat(szTitle, "")
    .ulFlags = BIF_RETURNONLYFSDIRS + BIF_DONTGOBELOWDOMAIN + BIF_STATUSTEXT
    .lpfnCallback = GetAddressofFunction(AddressOf BrowseCallbackProc)  'get address of function.
  End With

  lpIDList = SHBrowseForFolder(tBrowseInfo)
  If (lpIDList) Then
    sBuffer = Space(MAX_PATH)
    SHGetPathFromIDList lpIDList, sBuffer
    sBuffer = left(sBuffer, InStr(sBuffer, vbNullChar) - 1)
    BrowseForFolder = sBuffer
  Else
    BrowseForFolder = ""
  End If
  
End Function
 
Private Function BrowseCallbackProc(ByVal hWnd As Long, ByVal uMsg As Long, ByVal lp As Long, ByVal pData As Long) As Long
  
  Dim lpIDList As Long
  Dim ret As Long
  Dim sBuffer As String
  
  On Error Resume Next  'Sugested by MS to prevent an error from
                        'propagating back into the calling process.
     
  Select Case uMsg
  
    Case BFFM_INITIALIZED
      Call SendMessage(hWnd, BFFM_SETSELECTION, 1, m_CurrentDirectory)
      
    Case BFFM_SELCHANGED
      sBuffer = Space(MAX_PATH)
      
      ret = SHGetPathFromIDList(lp, sBuffer)
      If ret = 1 Then
        Call SendMessage(hWnd, BFFM_SETSTATUSTEXT, 0, sBuffer)
      End If
      
  End Select
  
  BrowseCallbackProc = 0
  
End Function

' This function allows you to assign a function pointer to a vaiable.
Private Function GetAddressofFunction(add As Long) As Long
  GetAddressofFunction = add
End Function

'===========================================================
' this function is to browse for a file in stead of a folder
'===========================================================
Function browsefile(myfile, mypath)
On Error GoTo myerror
   ' Requires reference to Microsoft Office 11.0 or higher Object Library.
   Dim fDialog As Office.FileDialog
   Dim varFile As Variant
   ' Set up the File Dialog.
   Set fDialog = Application.FileDialog(msoFileDialogFilePicker)
   With fDialog
      ' Allow user to not make multiple selections in dialog box
      .AllowMultiSelect = False
      ' Set the title of the dialog box.
      .title = "Please select the file to use"
      ' Clear out the current filters, and add our own.
      .Filters.Clear
      '.Filters.add "dbf", "*.dbf"
      ' Show the dialog box. If the .Show method returns True, the
      ' user picked at least one file. If the .Show method returns
      ' False, the user clicked Cancel.
      .InitialFileName = myfile
      If .show = True Then
         'Loop through each file selected and add it to our list box.
         For Each varFile In .SelectedItems
            browsefile = varFile
            myfile = (Mid$(browsefile, InStrRev(browsefile, "\") + 1))
            mypath = (Mid$(browsefile, 1, InStrRev(browsefile, "\") - 1))
         Next
      Else
         browsefile = False
      End If
   End With
   Exit Function
myerror:
    MsgBox "An error occurred; if this is the first time running this routine, check" & _
        "if there is a reference to the MS Office Object Library"
    browsefile = False
End Function



