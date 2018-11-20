VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   5205
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   13665
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5205
   ScaleWidth      =   13665
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H00EEEEEE&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1455
      Left            =   0
      ScaleHeight     =   1455
      ScaleWidth      =   13665
      TabIndex        =   0
      Top             =   0
      Width           =   13665
      Begin VB.Label lbl1 
         Alignment       =   2  'Center
         Caption         =   "New Project"
         Height          =   255
         Left            =   240
         TabIndex        =   1
         Top             =   960
         Width           =   1095
      End
      Begin VB.Image Image1 
         Appearance      =   0  'Flat
         Height          =   855
         Left            =   240
         Picture         =   "Form1.frx":0000
         Stretch         =   -1  'True
         Top             =   120
         Width           =   1095
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
End Sub

Private Sub Image1_Click()
    Image1.BorderStyle = BorderStyleConstants.vbBSSolid
End Sub
