VERSION 5.00
Begin VB.Form frmNewSnippet 
   Caption         =   "New Snippet"
   ClientHeight    =   5310
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   7185
   LinkTopic       =   "Form1"
   ScaleHeight     =   5310
   ScaleWidth      =   7185
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtSniName 
      Height          =   435
      Left            =   1890
      TabIndex        =   0
      Top             =   480
      Width           =   4650
   End
   Begin VB.TextBox txtDesc 
      Height          =   1605
      Left            =   1890
      MultiLine       =   -1  'True
      TabIndex        =   1
      Top             =   1320
      Width           =   4650
   End
   Begin VB.TextBox txtType 
      Height          =   435
      Left            =   1890
      MultiLine       =   -1  'True
      TabIndex        =   2
      Top             =   3360
      Width           =   4650
   End
   Begin VB.CommandButton cmdCreate 
      Caption         =   "Create Project"
      Default         =   -1  'True
      Height          =   435
      Left            =   4620
      TabIndex        =   3
      Top             =   4350
      Width           =   1935
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   450
      Left            =   285
      TabIndex        =   5
      Top             =   4335
      Width           =   1935
   End
   Begin VB.Label lblSniName 
      Caption         =   "Project Name"
      Height          =   405
      Index           =   0
      Left            =   330
      TabIndex        =   7
      Top             =   510
      Width           =   1215
   End
   Begin VB.Label lblDesc 
      Caption         =   "Description"
      Height          =   405
      Left            =   330
      TabIndex        =   6
      Top             =   1350
      Width           =   1215
   End
   Begin VB.Label lblType 
      Caption         =   "Type/Language"
      Height          =   375
      Left            =   330
      TabIndex        =   4
      Top             =   3360
      Width           =   1335
   End
End
Attribute VB_Name = "frmNewSnippet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim newSnippet As Snippet
Dim noUnloadPrompt As Boolean


Private Sub Form_Load()
    
    
    noUnloadPrompt = False
    
End Sub

Private Sub cmdCancel_Click()
    Me.Hide
    Unload Me
End Sub

Private Sub cmdCreate_Click()
    
    If checkFields = False Then
        Exit Sub
    End If
    Dim filName As String
    
    filName = GetAppPath(ssfLOCALAPPDATA)
    
    filName = filName & "\SourceTrace"
    
    filName = filName & "\Snippets\"
    
    CreateFolder (filName)
    
    filName = filName & CreateUID() & ".txt"
    
    Set newSnippet = New Snippet
    newSnippet.AddNew
    newSnippet.SnippetName = txtSniName.Text
    newSnippet.Description = txtDesc.Text
    newSnippet.SnippetType = txtType.Text
    
    newSnippet.Location = filName
    newSnippet.CFClass = "CF_" & UCase(txtSniName.Text)

    SetTopMostWindow frmNewSnippet.hwnd, False
    If newSnippet.SaveChanges() = 1 Then
        noUnloadPrompt = True
        
        'Go to snippet view
        frmSnippetView.SnippetId = newSnippet.SnippetId
        frmMDIMain.hideAllWindows
        frmSnippetView.Show
        frmSnippetView.WindowState = 2
        
        Unload Me
        Exit Sub
    Else
        MsgBox "Could not save changes. Please try again", vbOKOnly
    End If
    SetTopMostWindow frmNewSnippet.hwnd, True
End Sub

'functions

Private Function checkFields() As Boolean
    checkFields = True
    If txtSniName.Text = "" Then
        checkFields = False
    ElseIf txtDesc.Text = "" Then
        checkFields = False
    ElseIf txtType.Text = "" Then
        checkFields = False
    'Mode omitted as it is the default ref for now
    End If
End Function

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    SetTopMostWindow frmNewSnippet.hwnd, False
    If noUnloadPrompt Then
        Exit Sub
    End If
    
    If MsgBox("Cancel?", vbYesNo, "New Snippet") = vbNo Then
        Cancel = 1
        frmNewSnippet.Show
        SetTopMostWindow frmNewSnippet.hwnd, True
    End If
    
End Sub

