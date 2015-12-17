VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Cmodule"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'===============================
'class to manage forms and bas modules(import, export and delete)
'-------------------------------
Option Explicit
Public dict As Object
Private mdtmCreated As Date
Private thisproject As VBIDE.VBProject
Private modules As VBIDE.VBComponents
Private module As VBIDE.vbcomponent
Private i As Integer
Private fso, f, file
Private types As Object

Private Sub Class_Initialize()
  mdtmCreated = Now
  Set dict = CreateObject("Scripting.Dictionary")
  Set types = CreateObject("Scripting.Dictionary")
  Set thisproject = ActiveWorkbook.VBProject
  Set modules = thisproject.VBComponents
  loadModules
End Sub
Private Sub Class_Terminate()
    'all is well when class is set to nothing...
End Sub

Private Function loadModules(Optional ctype = 3)
'load vb attach types for referencing; we do not use vbext_ct_Document=100 ->worksheet code
'since they cannot be removed anyhow...
    types.add 1, ".bas" 'vbext_ct_StdModule=1
    types.add 2, ".cls" 'vbext_ct_ClassModule=2
    types.add 3, ".frm" 'vbext_ct_MSForm=3
'load all attached code with type ctype or smaller into the dictionary
    i = 0 'index for the dictionary
    For Each module In modules
        If (module.Type <= ctype) Then
               dict.add module.Name, module.Type
               i = i + 1
        End If
    Next
End Function
'----------------------------------
Public Function import(fldr)
'public function to import vba modules (file type<=ctype)from directory 'fldr'
    Dim replace As Boolean
        import = False
        Set fso = CreateObject("Scripting.FileSystemObject")
        On Error GoTo myerror 'in case folder does not exist
        Set f = fso.GetFolder(fldr)
        On Error GoTo 0 'reset
        If f.Files.Count = 0 Then
            Exit Function
        End If
        For Each file In f.Files
        'eliminate:
        ' - cmodule class (this module); it is part of the shell copy
        ' - .frx files; they are not compatible
            If Not file.Name = "Cmodule.cls" And Not Right(file.Name, 3) = "frx" Then
            'check if the module exists and make sure overwriting is okay
                Debug.Print file.Name
                If dict(left(file.Name, Len(file.Name) - 4)) Then
                    If replace = False Then
                        If MsgBox("Replace existing modules with the same name?", _
                            vbOKCancel) = vbCancel Then Exit Function
                        replace = True
                    End If
                    If replace = True Then
                        Set module = modules(left(file.Name, Len(file.Name) - 4))
                        modules.Remove module
                    End If
                End If
                modules.import fldr & "\" & file.Name
            End If
        Next
        import = True
        Exit Function
myerror:
        MsgBox "Sorry; I could not find any modules to import." & _
        "Either the folder: " & fldr & vbCr & _
        "does not exist or the folder has no modules:("
End Function

Public Function export(fldr, Optional ctype = 3)
'public function to export vba modules (file type<=ctype)from directory 'fldr'
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not dir(fldr, vbDirectory) = vbNullString Then
        If MsgBox("This will delete the existing modules...", vbOKCancel) = vbCancel Then
            export = False
            Exit Function
        End If
        fso.deletefolder (fldr)
    End If
    fso.createFolder (fldr)
    For Each module In modules
        If module.Type <= ctype Then _
           module.export fldr & "\" & module.Name & types(module.Type)
    Next
    export = True
End Function

Public Function delete(Optional ctype = 1)
    For Each module In modules
        If module.Type = ctype Then modules.Remove module
    Next
End Function
