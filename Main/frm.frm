VERSION 5.00
Begin VB.Form frm 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   5565
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11145
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5565
   ScaleWidth      =   11145
   ShowInTaskbar   =   0   'False
End
Attribute VB_Name = "frm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim mytask As New Task
Dim tasks As New TaskManager


Private Sub Form_Load()
    InitializeConnection
    tasks.IsDebug = True
    tasks.loadTasksFromDB
    
End Sub
