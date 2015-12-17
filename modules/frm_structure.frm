VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frm_structure 
   Caption         =   "Structuring your SAM"
   ClientHeight    =   3816
   ClientLeft      =   48
   ClientTop       =   372
   ClientWidth     =   3600
   OleObjectBlob   =   "frm_structure.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frm_structure"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

















Private Sub cmd_Click()
    myname = frm_structure.follow
    Application.Run myname
    Application.ScreenUpdating = True
End Sub

Private Sub Label_Click()

End Sub
