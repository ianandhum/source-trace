VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmNewProject 
   Caption         =   "New Project"
   ClientHeight    =   7170
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   7215
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7170
   ScaleWidth      =   7215
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog cdlLocation 
      Left            =   120
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdAddTag 
      Caption         =   "Add"
      Height          =   435
      Left            =   4560
      TabIndex        =   6
      Top             =   3720
      Width           =   855
   End
   Begin VB.CommandButton cmdClearTags 
      Caption         =   "Clear All"
      Height          =   435
      Left            =   5520
      TabIndex        =   7
      Top             =   3720
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   495
      Left            =   360
      TabIndex        =   14
      Top             =   6240
      Width           =   1935
   End
   Begin VB.CommandButton cmdCreate 
      Caption         =   "Create Project"
      Default         =   -1  'True
      Height          =   495
      Left            =   4680
      TabIndex        =   13
      Top             =   6240
      Width           =   1935
   End
   Begin VB.CommandButton cmdSelLoc 
      Caption         =   "Select Location"
      Height          =   435
      Left            =   5160
      TabIndex        =   12
      Top             =   5400
      Width           =   1455
   End
   Begin VB.TextBox txtLoc 
      Enabled         =   0   'False
      Height          =   435
      Left            =   1920
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   5400
      Width           =   3210
   End
   Begin VB.ComboBox cmbMode 
      Height          =   315
      IntegralHeight  =   0   'False
      ItemData        =   "frmNewProject.frx":0000
      Left            =   1920
      List            =   "frmNewProject.frx":000A
      TabIndex        =   9
      Text            =   "-- Select Mode --"
      Top             =   4920
      Width           =   4695
   End
   Begin VB.TextBox txtTags 
      Height          =   435
      Left            =   1920
      MultiLine       =   -1  'True
      TabIndex        =   5
      Top             =   3720
      Width           =   2610
   End
   Begin VB.TextBox txtDesc 
      Height          =   1605
      Left            =   1920
      TabIndex        =   2
      Top             =   1680
      Width           =   4650
   End
   Begin VB.TextBox txtPrjName 
      Height          =   435
      Left            =   1920
      TabIndex        =   1
      Top             =   840
      Width           =   4650
   End
   Begin VB.Label lblSelTags 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1920
      TabIndex        =   15
      Top             =   4200
      Width           =   4695
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label1 
      Caption         =   "Location"
      Height          =   375
      Left            =   360
      TabIndex        =   10
      Top             =   5520
      Width           =   1335
   End
   Begin VB.Label lblMode 
      Caption         =   "Mode"
      Height          =   375
      Left            =   360
      TabIndex        =   8
      Top             =   4920
      Width           =   1335
   End
   Begin VB.Label lblTags 
      Caption         =   "Tags"
      Height          =   375
      Left            =   360
      TabIndex        =   4
      Top             =   3720
      Width           =   1335
   End
   Begin VB.Label lblDesc 
      Caption         =   "Description"
      Height          =   405
      Left            =   360
      TabIndex        =   3
      Top             =   1710
      Width           =   1215
   End
   Begin VB.Label lblPrjName 
      Caption         =   "Project Name"
      Height          =   405
      Index           =   0
      Left            =   360
      TabIndex        =   0
      Top             =   870
      Width           =   1215
   End
End
Attribute VB_Name = "frmNewProject"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim newPrj As Project
Dim noUnloadPrompt As Boolean


Private Sub Form_Load()
    ' we have ref mode only right now
    cmbMode.ListIndex = 0
    cmbMode.Enabled = False
    noUnloadPrompt = False
    
End Sub

Private Sub cmdCancel_Click()
    Me.Hide
    Unload Me
End Sub

Private Sub cmdCreate_Click()
    
    If checkFields = False Then
        
        SetTopMostWindow frmNewProject.hwnd, False
        MsgBox "All fields are mandatory", vbOKOnly
        SetTopMostWindow frmNewProject.hwnd, True

        Exit Sub
    
    End If
    
    Set newPrj = New Project
    newPrj.AddNew
    newPrj.ProjectName = txtPrjName.Text
    newPrj.Description = txtDesc.Text
    'Only ref mode for now
    newPrj.Mode = "REF"
    newPrj.Tags = lblSelTags.Caption
    newPrj.Location = txtLoc.Text
    newPrj.CFClass = "CF_" & UCase(txtPrjName.Text)
    newPrj.State = "ACTIVE"
    newPrj.LastUpdate = DateTime.Now

    SetTopMostWindow frmNewProject.hwnd, False
    If newPrj.SaveChanges() = 1 Then
        noUnloadPrompt = True
        Unload Me
        MsgBox "Project Created Successfully", vbOKOnly
        Exit Sub
    Else
        MsgBox "Could not save changes. Please try again", vbOKOnly
    End If
    SetTopMostWindow frmNewProject.hwnd, True
End Sub

Private Sub cmdSelLoc_Click()
    Dim sTempDir As String
    On Error Resume Next
    sTempDir = CurDir
    cdlLocation.DialogTitle = "Select a directory"
    cdlLocation.InitDir = App.path
    cdlLocation.FileName = "Select a Directory"
    cdlLocation.FLAGS = cdlOFNNoValidate + cdlOFNHideReadOnly
    cdlLocation.Filter = "Directories|*.~#~"
    cdlLocation.CancelError = True
    cdlLocation.ShowOpen

    If Err <> 32755 Then    ' if not cancel
        Me.txtLoc.Text = CurDir
    End If
    ChDir sTempDir
End Sub

Private Sub cmdAddTag_Click()
addTag
End Sub

Private Sub cmdClearTags_Click()
    lblSelTags.Caption = ""
    txtTags.Text = ""
End Sub



'functions

Private Sub addTag()
    Dim tagStr$
    
    tagStr = lblSelTags.Caption
    If tagStr = "" Then
        lblSelTags.Caption = txtTags.Text
    Else
        lblSelTags.Caption = lblSelTags.Caption & "," & txtTags.Text
    
    End If
    
    txtTags.Text = ""
    txtTags.SetFocus
End Sub

Private Function checkFields() As Boolean
    checkFields = True
    If txtPrjName.Text = "" Then
        checkFields = False
    ElseIf txtDesc.Text = "" Then
        checkFields = False
    ElseIf txtLoc.Text = "" Then
        checkFields = False
    ElseIf lblSelTags.Caption = "" Then
        checkFields = False
    'Mode omitted as it is the default ref for now
    End If
End Function

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    SetTopMostWindow frmNewProject.hwnd, False
    If noUnloadPrompt Then
        Exit Sub
    End If
    
    If MsgBox("Project will not be created if you choose Yes. Are you Sure?", vbYesNo, "New Project") = vbNo Then
        Cancel = 1
        frmNewProject.Show
        SetTopMostWindow frmNewProject.hwnd, True
    End If
    
End Sub
