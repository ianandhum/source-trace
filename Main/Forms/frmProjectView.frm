VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmProjectView 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "None"
   ClientHeight    =   9840
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   20490
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9840
   ScaleWidth      =   20490
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox pbxNavTiles 
      Appearance      =   0  'Flat
      BackColor       =   &H00EEEE0E&
      BorderStyle     =   0  'None
      ForeColor       =   &H8000000B&
      Height          =   1275
      Index           =   1
      Left            =   2310
      ScaleHeight     =   1275
      ScaleWidth      =   1605
      TabIndex        =   14
      Top             =   90
      Width           =   1605
      Begin VB.Label lblNavTile 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "label_text_here"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   1
         Left            =   60
         TabIndex        =   15
         Top             =   915
         Width           =   1500
      End
      Begin VB.Image imgNavTile 
         Height          =   855
         Index           =   1
         Left            =   225
         Picture         =   "frmProjectView.frx":0000
         Stretch         =   -1  'True
         Top             =   15
         Width           =   1155
      End
   End
   Begin VB.CommandButton btnFrmStats 
      Appearance      =   0  'Flat
      Caption         =   "History"
      Enabled         =   0   'False
      Height          =   495
      Index           =   2
      Left            =   7575
      TabIndex        =   11
      Top             =   7560
      Width           =   2295
   End
   Begin VB.CommandButton btnFrmStats 
      Appearance      =   0  'Flat
      Caption         =   "Log"
      Height          =   495
      Index           =   1
      Left            =   5295
      TabIndex        =   10
      Top             =   7560
      Width           =   2295
   End
   Begin VB.CommandButton btnFrmStats 
      Appearance      =   0  'Flat
      Caption         =   "Commits"
      Height          =   495
      Index           =   0
      Left            =   3015
      TabIndex        =   9
      Top             =   7560
      Width           =   2295
   End
   Begin VB.PictureBox pbxStats 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1935
      Left            =   3010
      ScaleHeight     =   1935
      ScaleWidth      =   17535
      TabIndex        =   5
      Top             =   8040
      Width           =   17535
      Begin VB.Frame frmStats 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1215
         Index           =   2
         Left            =   1200
         TabIndex        =   8
         Top             =   120
         Width           =   5895
      End
      Begin VB.Frame frmStats 
         Appearance      =   0  'Flat
         BackColor       =   &H80000002&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1215
         Index           =   1
         Left            =   4320
         TabIndex        =   7
         Top             =   600
         Width           =   5895
      End
      Begin VB.Frame frmStats 
         Appearance      =   0  'Flat
         BackColor       =   &H80000003&
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         ForeColor       =   &H80000008&
         Height          =   1215
         Index           =   0
         Left            =   0
         TabIndex        =   6
         Top             =   0
         Width           =   5895
      End
   End
   Begin VB.PictureBox pbxHeader 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H00EEEEEE&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1455
      Left            =   0
      ScaleHeight     =   1455
      ScaleWidth      =   20490
      TabIndex        =   0
      Top             =   0
      Width           =   20490
      Begin VB.PictureBox pbxNavTiles 
         Appearance      =   0  'Flat
         BackColor       =   &H00EEEE0E&
         BorderStyle     =   0  'None
         ForeColor       =   &H8000000B&
         Height          =   1275
         Index           =   0
         Left            =   480
         ScaleHeight     =   1275
         ScaleWidth      =   1605
         TabIndex        =   12
         Top             =   90
         Width           =   1605
         Begin VB.Image imgNavTile 
            Height          =   855
            Index           =   0
            Left            =   255
            Picture         =   "frmProjectView.frx":7F03
            Stretch         =   -1  'True
            Top             =   15
            Width           =   1155
         End
         Begin VB.Label lblNavTile 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "label_text_here"
            BeginProperty Font 
               Name            =   "Segoe UI"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   0
            Left            =   60
            TabIndex        =   13
            Top             =   915
            Width           =   1500
         End
      End
   End
   Begin ComctlLib.TreeView tvFileNodes 
      Height          =   7815
      Left            =   0
      TabIndex        =   1
      Top             =   1920
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   13785
      _Version        =   327682
      LabelEdit       =   1
      LineStyle       =   1
      Style           =   7
      Appearance      =   0
      MousePointer    =   4
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Consolas"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin RichTextLib.RichTextBox rtbFileView 
      Height          =   5655
      Left            =   3010
      TabIndex        =   2
      Top             =   1920
      Width           =   17520
      _ExtentX        =   30903
      _ExtentY        =   9975
      _Version        =   393217
      BackColor       =   16448250
      BorderStyle     =   0
      ScrollBars      =   3
      Appearance      =   0
      TextRTF         =   $"frmProjectView.frx":FE06
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Consolas"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Shape shpHeaderLine 
      BackColor       =   &H00FFFFFF&
      BorderColor     =   &H00808080&
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   495
      Index           =   2
      Left            =   3010
      Top             =   7560
      Width           =   17535
   End
   Begin VB.Line lnHeader 
      BorderColor     =   &H00C0C0C0&
      X1              =   3000
      X2              =   3000
      Y1              =   1440
      Y2              =   9840
   End
   Begin VB.Label lblHeaderLineText 
      BackStyle       =   0  'Transparent
      Caption         =   "file name"
      BeginProperty Font 
         Name            =   "Segoe UI Semibold"
         Size            =   9.75
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   3240
      TabIndex        =   4
      Top             =   1545
      Width           =   5655
   End
   Begin VB.Label lblHeaderLineText 
      BackStyle       =   0  'Transparent
      Caption         =   "ProjectName"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   120
      TabIndex        =   3
      Top             =   1550
      Width           =   2895
   End
   Begin VB.Shape shpHeaderLine 
      BackColor       =   &H00FFFFFF&
      BorderColor     =   &H00808080&
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   495
      Index           =   1
      Left            =   3000
      Top             =   1440
      Width           =   17535
   End
   Begin VB.Shape shpHeaderLine 
      BorderColor     =   &H00808080&
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   495
      Index           =   0
      Left            =   0
      Top             =   1440
      Width           =   3015
   End
End
Attribute VB_Name = "frmProjectView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private InitialControlList() As ControlInitial

Private Sub Form_Load()

    InitialControlList = GetLocation(Me)
    ReSizePosForm Me, Me.height, Me.width, Me.Left, Me.Top, True
    ResizeShapes
    initNavTiles
    
    
End Sub

Private Sub Form_Resize()
    ResizeControls Me, InitialControlList, True
    ResizeShapes
End Sub
Private Sub ResizeShapes()
    
    lnHeader.Y1 = shpHeaderLine(0).Top
    lnHeader.Y2 = frmProjectView.height
    lnHeader.X1 = tvFileNodes.Left + tvFileNodes.width + 100
    lnHeader.X2 = tvFileNodes.Left + tvFileNodes.width + 100
    pbxStats.Top = shpHeaderLine(2).Top + shpHeaderLine(2).height - 2
    pbxStats.Left = tvFileNodes.Left + tvFileNodes.width + 120
    pbxStats.height = Me.height - shpHeaderLine(2).Top
    For i = 0 To frmStats.Count - 1
        frmStats(i).height = pbxStats.height
        frmStats(i).width = pbxStats.width
        frmStats(i).Left = 0
        frmStats(i).Top = 0
        
    Next i
    
    
    
End Sub

Private Sub initNavTiles()
    For i = 0 To pbxNavTiles.Count - 1
        pbxNavTiles(i).BackColor = &HEEEEEE
        pbxNavTiles(i).BorderStyle = BorderStyleConstants.vbTransparent
    Next i
    
End Sub

Private Sub imgNavTile_Click(Index As Integer)
    MsgBox Index
    Me.Hide
    frm.Show
    frm.WindowState = 2
    
End Sub

Private Sub imgNavTile_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    pbxNavTiles(Index).BorderStyle = BorderStyleConstants.vbBSSolid
    For i = 0 To pbxNavTiles.Count - 1
        If Index <> i Then
            pbxNavTiles(i).BackColor = &HEEEEEE
            pbxNavTiles(i).BorderStyle = BorderStyleConstants.vbTransparent
        End If
    Next i
    
    
    
End Sub
