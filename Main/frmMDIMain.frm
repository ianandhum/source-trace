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
      Begin VB.Menu mnu_new 
         Caption         =   "New"
         Begin VB.Menu mnu_file_new_prj 
            Caption         =   "Project"
         End
         Begin VB.Menu mnu_new_sni 
            Caption         =   "Snippet"
         End
      End
      Begin VB.Menu mnu_quit 
         Caption         =   "Quit"
      End
   End
   Begin VB.Menu mnu_top_view 
      Caption         =   "View"
      Index           =   2
      Begin VB.Menu mnu_view_overview 
         Caption         =   "Overview"
      End
      Begin VB.Menu mnu_view_Projects 
         Caption         =   "Projects"
      End
      Begin VB.Menu mnu_view_tasks 
         Caption         =   "Tasks"
      End
      Begin VB.Menu mnu_view_snippet 
         Caption         =   "Snippets"
      End
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
    Call changeMDIView(frmOverView)
End Sub

Private Sub MDIForm_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    End
End Sub

Private Sub mnu_file_new_prj_Click()
    frmNewProject.Show
    SetTopMostWindow frmNewProject.hwnd, True
End Sub

Private Sub mnu_new_sni_Click()
    frmNewSnippet.Show
    SetTopMostWindow frmNewSnippet.hwnd, True
End Sub

Private Sub mnu_quit_Click()
    End
End Sub

Private Sub mnu_view_overview_Click()
    
    Call changeMDIView(frmOverView)
End Sub

Private Sub mnu_view_Projects_Click()
    
    Call changeMDIView(frmProjectList)
End Sub

Private Sub mnu_view_snippet_Click()
    
    Call changeMDIView(frmSnippetList)
End Sub

Private Sub mnu_view_tasks_Click()
    Call changeMDIView(frmTaskView)
End Sub

Public Sub changeMDIView(Source As Form)
    
    hideAllWindows
    Source.Show
    Source.WindowState = 2
End Sub

Public Sub hideAllWindows()
On Error Resume Next
    Unload frmNewProject
    Unload frmProjectView
    Unload frmTaskView
    Unload frmSnippetView
    Unload frmOverView
End Sub

Public Sub filterMenu(mnuList() As Integer)
    
End Sub
