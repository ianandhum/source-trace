VERSION 5.00
Begin VB.MDIForm frmMDIMain 
   Appearance      =   0  'Flat
   BackColor       =   &H00FFFFFF&
   Caption         =   "Source Trace - Projects"
   ClientHeight    =   5985
   ClientLeft      =   3945
   ClientTop       =   3090
   ClientWidth     =   13590
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.Menu mnu_top_file 
      Caption         =   "File"
      Index           =   0
   End
   Begin VB.Menu mnu_top_edit 
      Caption         =   "Edit"
      Index           =   1
   End
   Begin VB.Menu mnu_top_view 
      Caption         =   "View"
      Index           =   2
   End
   Begin VB.Menu mnu_top_repo 
      Caption         =   "Repository"
      Index           =   3
   End
   Begin VB.Menu mnu_top_tools 
      Caption         =   "Tools"
      Index           =   4
   End
   Begin VB.Menu mnu_top_help 
      Caption         =   "Help"
      Index           =   5
   End
End
Attribute VB_Name = "frmMDIMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub MDIForm_Load()

    frmProjectView.Show
    frmProjectView.WindowState = 2
    
End Sub
