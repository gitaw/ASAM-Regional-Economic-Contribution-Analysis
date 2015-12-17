VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frm_sam 
   Caption         =   "SAM structure"
   ClientHeight    =   5040
   ClientLeft      =   48
   ClientTop       =   372
   ClientWidth     =   4560
   OleObjectBlob   =   "frm_sam.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frm_sam"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

















Private Sub cmd_Click()
    js = frm_sam.samend.Value
    je = frm_sam.endogenous.Value
    jc = frm_sam.sectors.Value
    jw = frm_sam.compensation.Value
    continuefromfsam
End Sub


Private Sub UserForm_Initialize()
    dollars.Value = "1,000"
    dollars.AddItem ("1")
End Sub
