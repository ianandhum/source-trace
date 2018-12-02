VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmNewTask 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Add New Task"
   ClientHeight    =   4095
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   5940
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4095
   ScaleWidth      =   5940
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin MSComCtl2.DTPicker dtNewTaskDueDate 
      Height          =   375
      Left            =   2040
      TabIndex        =   7
      Top             =   2520
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   661
      _Version        =   393216
      CalendarBackColor=   16777215
      Format          =   160497665
      CurrentDate     =   43434
   End
   Begin VB.TextBox txtNewTaskDesc 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   2040
      MultiLine       =   -1  'True
      TabIndex        =   4
      Top             =   1440
      Width           =   3615
   End
   Begin VB.TextBox txtNewTask 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2040
      TabIndex        =   2
      Top             =   720
      Width           =   3615
   End
   Begin VB.CommandButton CancelButton 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3120
      TabIndex        =   1
      Top             =   3480
      Width           =   1215
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   4440
      TabIndex        =   0
      Top             =   3480
      Width           =   1215
   End
   Begin VB.Label lblNewTaskDueDate 
      Caption         =   "Due Date"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   2640
      Width           =   1815
   End
   Begin VB.Label lblNewTaskDesc 
      Caption         =   "Description"
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   1440
      Width           =   1815
   End
   Begin VB.Label lblNewTask 
      Caption         =   "Task Name"
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   720
      Width           =   1815
   End
End
Attribute VB_Name = "frmNewTask"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Public CardIndex As Integer
Private Sub CancelButton_Click()
    Unload Me
End Sub

Private Sub Form_Activate()
    txtNewTask.SetFocus
    dtNewTaskDueDate.value = DateTime.Date
End Sub

Private Sub Form_Load()
    SetTopMostWindow Me.hwnd, True
End Sub

Private Sub OKButton_Click()
    If txtNewTask.Text = "" Or txtNewTaskDesc.Text = "" Or dtNewTaskDueDate.value = "" Then
        Exit Sub
    End If
    Call frmTaskView.addNewTile(CardIndex, txtNewTask.Text, txtNewTaskDesc.Text, dtNewTaskDueDate.value)
    SetTopMostWindow Me.hwnd, True
    Unload Me
End Sub

