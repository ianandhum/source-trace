VERSION 5.00
Begin VB.MDIForm frmMDIMain 
   Appearance      =   0  'Flat
   BackColor       =   &H00FFFFFF&
   Caption         =   "Source Trace"
   ClientHeight    =   5985
   ClientLeft      =   165
   ClientTop       =   810
   ClientWidth     =   13590
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Menu mnu_top_file 
      Caption         =   "File"
      Index           =   0
      Begin VB.Menu mnu_file_new_prj 
         Caption         =   "New Project"
      End
      Begin VB.Menu mnu_quit 
         Caption         =   "Quit"
      End
   End
   Begin VB.Menu mnu_top_edit 
      Caption         =   "Edit"
      Index           =   1
   End
   Begin VB.Menu mnu_top_view 
      Caption         =   "View"
      Index           =   2
      Begin VB.Menu mnu_view_Projects 
         Caption         =   "Projects"
      End
      Begin VB.Menu mnu_view_tasks 
         Caption         =   "Tasks"
      End
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
Dim cmd As New CmdRunner

Private Sub MDIForm_Load()
    InitializeConnection
End Sub

Private Sub MDIForm_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    End
End Sub

Private Sub mnu_file_new_prj_Click()
    frmNewProject.Show
    SetTopMostWindow frmNewProject.hwnd, True
End Sub



Private Sub mnu_quit_Click()
    End
End Sub

Private Sub mnu_view_Projects_Click()
    hideAllWindows
    frmProjectView.Show
    frmProjectView.WindowState = 2
End Sub

Private Sub mnu_view_tasks_Click()
    hideAllWindows
    frmTaskView.Show
    frmTaskView.WindowState = 2
End Sub

Private Sub hideAllWindows()
On Error Resume Next
    Unload frmNewProject
    Unload frmProjectView
    Unload frmTaskView
End Sub
