VERSION 5.00
Begin VB.Form frmProjectList 
   BorderStyle     =   0  'None
   Caption         =   "Projects"
   ClientHeight    =   8385
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   19050
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8385
   ScaleWidth      =   19050
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox pbxHead 
      Appearance      =   0  'Flat
      BackColor       =   &H00FCFCFC&
      CausesValidation=   0   'False
      ForeColor       =   &H80000008&
      Height          =   1455
      Left            =   0
      ScaleHeight     =   1425
      ScaleWidth      =   18840
      TabIndex        =   8
      Top             =   0
      Width           =   18870
      Begin VB.PictureBox pbxOption 
         BackColor       =   &H00FCFCFC&
         BorderStyle     =   0  'None
         Height          =   1215
         Index           =   0
         Left            =   17295
         ScaleHeight     =   1215
         ScaleWidth      =   1230
         TabIndex        =   9
         Top             =   165
         Width           =   1230
         Begin VB.Label lblOptionIcon 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "+"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   24
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   495
            Index           =   0
            Left            =   405
            TabIndex        =   11
            Top             =   195
            Width           =   495
         End
         Begin VB.Label lblOption 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "New"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   0
            Left            =   45
            TabIndex        =   10
            Top             =   870
            Width           =   1215
         End
         Begin VB.Shape shpOptionIcon 
            BackColor       =   &H00E0E0E0&
            BackStyle       =   1  'Opaque
            BorderStyle     =   0  'Transparent
            FillColor       =   &H008080FF&
            Height          =   735
            Index           =   0
            Left            =   165
            Shape           =   3  'Circle
            Top             =   90
            Width           =   975
         End
      End
      Begin VB.Label lblProjectHeader 
         BackStyle       =   0  'Transparent
         Caption         =   "Projects"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   270
         TabIndex        =   13
         Top             =   330
         Width           =   6375
      End
      Begin VB.Label lblHeadDesc 
         BackStyle       =   0  'Transparent
         Caption         =   "These are Projects created using Project Wizard"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   540
         Left            =   285
         TabIndex        =   12
         Top             =   735
         Width           =   7830
      End
   End
   Begin VB.PictureBox pbxContent 
      BorderStyle     =   0  'None
      Height          =   8070
      Left            =   180
      ScaleHeight     =   8070
      ScaleWidth      =   14325
      TabIndex        =   0
      Top             =   1440
      Width           =   14325
      Begin VB.PictureBox pbxContainer 
         Appearance      =   0  'Flat
         BackColor       =   &H00EFEFEF&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   4350
         Left            =   180
         ScaleHeight     =   4350
         ScaleWidth      =   11655
         TabIndex        =   2
         Top             =   1335
         Width           =   11655
         Begin VB.PictureBox pbxProjectTile 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   1320
            Index           =   0
            Left            =   450
            ScaleHeight     =   1320
            ScaleWidth      =   7965
            TabIndex        =   3
            Top             =   2505
            Visible         =   0   'False
            Width           =   7965
            Begin VB.Label lblProjectDesc 
               BackStyle       =   0  'Transparent
               Caption         =   "Snippet Description"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   585
               Index           =   0
               Left            =   210
               TabIndex        =   6
               Top             =   615
               Width           =   7575
            End
            Begin VB.Label lblProjectName 
               BackStyle       =   0  'Transparent
               Caption         =   "Snippet Title"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   435
               Index           =   0
               Left            =   210
               TabIndex        =   5
               Top             =   270
               Width           =   4530
            End
            Begin VB.Label lblProjectView 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "View Project"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   -1  'True
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Index           =   0
               Left            =   6705
               TabIndex        =   4
               Top             =   180
               Width           =   1095
            End
         End
         Begin VB.Label lblTypeHeader 
            BackStyle       =   0  'Transparent
            Caption         =   "SNIPPETS"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   375
            Index           =   0
            Left            =   495
            TabIndex        =   7
            Top             =   1845
            Width           =   6375
         End
      End
      Begin VB.VScrollBar vsrContainer 
         Height          =   3135
         Left            =   12060
         SmallChange     =   200
         TabIndex        =   1
         Top             =   1200
         Visible         =   0   'False
         Width           =   225
      End
   End
End
Attribute VB_Name = "frmProjectList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim prjManager As ProjectManager
Dim tmpProject As Project
Dim loaded As Boolean
Private Sub Form_Load()
    adjustContainers
    'initProjectList
    pbxProjectTile(0).Left = -(pbxProjectTile(0).width)
    
End Sub
Private Sub Form_Resize()
    Set prjManager = Nothing
    If Me.WindowState = 2 Then
        adjustContainers
        alignControls
        initProjectList
        pbxProjectTile(0).Left = -(pbxProjectTile(0).width)
    End If
    
End Sub






Private Sub lblProjectView_Click(Index As Integer)
'Main Navigation
    frmMDIMain.hideAllWindows
    frmProjectView.ProjectId = Val(pbxProjectTile(Index).tag)
    frmProjectView.Show
    frmProjectView.WindowState = 2
End Sub

Private Sub lblProjectView_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call changeProjectsStyle(Index, &HFFDDBB)
    
End Sub

Private Sub lblProjectView_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call changeProjectsStyle(Index, vbWhite)
End Sub

Private Sub lblProjectView_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    SetCursor LoadCursor(0, IDC_HAND)
End Sub

Private Sub vsrContainer_Change()
    Dim changeVal As Long
    changeVal = -vsrContainer.value
    
    If moveTop < pbxContainer.height Then
        pbxContainer.Top = changeVal& * 10
    End If
End Sub



Private Sub changeProjectsStyle(Index As Integer, tBackColor As OLE_COLOR)
    pbxProjectTile(Index).BackColor = tBackColor
End Sub


Public Sub CheckKeyCode(KeyCode As Integer)

    Dim nScrollValue As Double
    Dim nOnePage As Integer
    
    nOnePage = Me.vsrContainer.height
    
    If KeyCode = vbKeyPageUp Or KeyCode = vbKeyPageDown Then
        If KeyCode = vbKeyPageDown Then
            nScrollValue = -Me.pbxContainer.Top + nOnePage
        Else
            nScrollValue = -Me.pbxContainer.Top - nOnePage
        End If
        If nScrollValue > Me.vsrContainer.Max Then
            nScrollValue = Me.vsrContainer.Max
            Me.pbxContainer.Top = -Me.vsrContainer.Max
        End If
        If nScrollValue > 0 Then
            Me.vsrContainer.value = nScrollValue
        Else
            Me.vsrContainer.value = 0
        End If
    End If
    
End Sub


Private Sub pbxContent_KeyDown(KeyCode As Integer, Shift As Integer)

    CheckKeyCode KeyCode
    
End Sub

Private Sub pbxContainer_KeyDown(KeyCode As Integer, Shift As Integer)

    CheckKeyCode KeyCode
    
End Sub




Private Sub pbxOption_Click(Index As Integer)
    
    frmNewProject.Show
    SetTopMostWindow frmNewProject.hwnd, True
End Sub

Private Sub lblOptionIcon_Click(Index As Integer)
    
    frmNewProject.Show
    SetTopMostWindow frmNewProject.hwnd, True
End Sub


Private Sub pbxOption_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    shpOptionIcon(Index).BackColor = &H217866
    lblOption(Index).ForeColor = &H217866
    
End Sub

Private Sub pbxOption_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    SetCursor LoadCursor(0, IDC_HAND)
    
End Sub

Private Sub pbxOption_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    shpOptionIcon(Index).BackColor = &HE0E0E0
    lblOption(Index).ForeColor = &H111111
    
End Sub

Private Sub lblOptionIcon_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    shpOptionIcon(Index).BackColor = &H217866
    lblOption(Index).ForeColor = &H217866
    
End Sub

Private Sub lblOptionIcon_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    SetCursor LoadCursor(0, IDC_HAND)
    
End Sub

Private Sub lblOptionIcon_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    shpOptionIcon(Index).BackColor = &HE0E0E0
    lblOption(Index).ForeColor = &H111111
End Sub












'Functions

Private Sub adjustContainers()
    pbxContent.height = Me.height - pbxHead.height
    pbxContent.width = Me.width
    pbxContent.Left = 0
    pbxContent.Top = pbxHead.height + 1
    
    pbxContainer.height = pbxContent.height
    pbxContainer.width = pbxContent.width
    pbxContainer.Left = 0
    pbxContainer.Top = 0
    
    pbxHead.width = Me.width + 10
    pbxHead.Top = -20
    pbxHead.Left = -10
    
    pbxOption(0).Left = pbxHead.width - pbxOption(0).width - 120
    
End Sub

Private Sub initProjectList()

        Set prjManager = New ProjectManager
        Call prjManager.loadProjectsFromDB(" 1=1 ORDER BY last_update DESC ")
        Call alignControls
        If prjManager.IsLoaded Then
            Dim Tags As String
            For i = 1 To prjManager.Count
                
                Tags = UCase(prjManager.Projects(i).Tags)
                If i > 1 Then
                    If UCase(prjManager.Projects(i - 1).Tags) = UCase(prjManager.Projects(i).Tags) Then
                        Tags = ""
                    End If
                End If
                If prjManager.Projects(i).IsLoaded Then
                    Call addProjectTile(i, prjManager.Projects(i).ProjectName, prjManager.Projects(i).Description, Tags, prjManager.Projects(i).ProjectId)
                End If
            Next i
        End If
End Sub

Private Sub addProjectTile(ByVal Index As Integer, ByVal name As String, ByVal desc As String, Tags As String, Optional tag As String = "")
    CreateProjectTile
    Call SetProjectTileInfo(Index, name, desc, tag)
    If Index > 1 Then
        Dim nextPos As Long
        nextPos = pbxProjectTile(Index - 1).Top + pbxProjectTile(Index - 1).height + 200
    
        If Tags <> "" Then
            Call CreateTypeHeader(lblTypeHeader.Count, Tags)
            Call AdjustTypeHeader(lblTypeHeader.Count - 1, nextPos)
            nextPos = nextPos + lblTypeHeader(lblTypeHeader.Count - 1).height + 200
        End If
    
        pbxProjectTile(Index).Top = nextPos
    Else
        lblTypeHeader(0).Caption = Tags
        lblTypeHeader(0).Visible = True
        Call AdjustTypeHeader(0, lblProjectHeader.Top + lblProjectHeader.height + 260)
        pbxProjectTile(1).Top = lblTypeHeader(0).height + lblTypeHeader(0).Top
        
    End If
    Call AdjustProjectTile(Index)
    Call AdjustContainer(Index)
    
End Sub
Private Sub CreateTypeHeader(Index As Integer, Tags As String)
    Load lblTypeHeader(Index)
    lblTypeHeader(Index).Caption = Tags
    lblTypeHeader(Index).Visible = True
    
            
End Sub

Private Function CreateProjectTile() As Integer
    Dim nextIndex As Integer
    nextIndex = pbxProjectTile.Count
    
    'load the contents
    Load pbxProjectTile(nextIndex)
    Load lblProjectName(nextIndex)
    Load lblProjectDesc(nextIndex)
    Load lblProjectView(nextIndex)
    
    Set lblProjectName(lblProjectName.Count - 1).Container = pbxProjectTile(nextIndex)
    Set lblProjectDesc(lblProjectDesc.Count - 1).Container = pbxProjectTile(nextIndex)
    Set lblProjectView(lblProjectView.Count - 1).Container = pbxProjectTile(nextIndex)
    
    pbxProjectTile(pbxProjectTile.Count - 1).Visible = True
    lblProjectName(lblProjectName.Count - 1).Visible = True
    lblProjectDesc(lblProjectDesc.Count - 1).Visible = True
    lblProjectView(lblProjectView.Count - 1).Visible = True
    
    CreateProjectTile = pbxProjectTile.Count - 1

End Function

Private Sub AdjustTypeHeader(Index As Integer, nextPos As Long)
    lblTypeHeader(Index).Top = nextPos + 50
    lblTypeHeader(Index).Left = pbxContainer.width * 0.1
End Sub

Private Sub AdjustProjectTile(Index As Integer)
    
    
    pbxProjectTile(Index).Left = pbxContainer.width * 0.1
    pbxProjectTile(Index).width = pbxContainer.width * 0.8
    
    lblProjectView(Index).Left = pbxProjectTile(Index).width - lblProjectView(Index).width - 100
    
End Sub


Private Sub SetProjectTileInfo(Index As Integer, name As String, desc As String, Optional tag As String = "")
    
    lblProjectName(Index).Caption = name
    lblProjectDesc(Index).Caption = desc
    
    If Len(tag) > 0 Then
        pbxProjectTile(Index).tag = tag
    End If

End Sub


Private Sub AdjustContainer(Index As Integer)
    
    pbxContainer.height = pbxProjectTile(Index).Top + pbxProjectTile(Index).height + 500
    
    Dim maxScroll As Double
    
    maxScroll = pbxContainer.height - pbxContent.height
    
    If pbxContainer.height > (pbxContent.height) Then
        vsrContainer.Max = maxScroll / 10
        vsrContainer.value = 0
        
        vsrContainer.LargeChange = pbxContainer.height / 10
        
        vsrContainer.Top = pbxContainer.Top
        vsrContainer.Left = pbxContainer.width - vsrContainer.width
        vsrContainer.height = pbxContent.height
        vsrContainer.Visible = True
    End If
End Sub

Private Sub clearControls()
    Dim i As Integer
    i = lblProjectName.Count - 1
    While i > 1
        Unload lblProjectName(i)
        i = i - 1
    Wend
    
    i = lblProjectDesc.Count - 1
    While i > 1
        Unload lblProjectDesc(i)
        i = i - 1
    Wend
    
    i = lblProjectView.Count - 1
    While i > 1
        Unload lblProjectView(i)
        i = i - 1
    Wend
    
    i = lblTypeHeader.Count - 1
    While i > 1
        Unload lblTypeHeader(i)
        i = i - 1
    Wend
    
    i = pbxProjectTile.Count
    While i > 1
        Unload pbxProjectTile(i)
        i = i - 1
    Wend
    
    
    
End Sub

Private Sub alignControls()
    AdjustContainer (pbxProjectTile.Count - 1)
End Sub

