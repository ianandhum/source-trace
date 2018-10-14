VERSION 5.00
Begin VB.MDIForm frmMDIMain 
   Appearance      =   0  'Flat
   BackColor       =   &H00FFFFFF&
   Caption         =   "MDIForm1"
   ClientHeight    =   5310
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   10530
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
End
Attribute VB_Name = "frmMDIMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub MDIForm_Load()
frm.Show
frm.WindowState = 2


End Sub
