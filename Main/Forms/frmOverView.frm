VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmOverView 
   BackColor       =   &H00F3F3F3&
   BorderStyle     =   0  'None
   Caption         =   "Home"
   ClientHeight    =   7950
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   16710
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7950
   ScaleWidth      =   16710
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox pbxTaskTile 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      FillColor       =   &H008080FF&
      ForeColor       =   &H008080FF&
      Height          =   1635
      Index           =   0
      Left            =   7140
      ScaleHeight     =   1635
      ScaleWidth      =   8025
      TabIndex        =   7
      Top             =   4065
      Visible         =   0   'False
      Width           =   8025
      Begin VB.Label lblTaskTileDesc 
         BackStyle       =   0  'Transparent
         Caption         =   "This is the description section can that is shown here"
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
         Height          =   555
         Index           =   0
         Left            =   1170
         TabIndex        =   13
         Top             =   540
         Width           =   3615
      End
      Begin VB.Label lblTaskTileProjectName 
         BackStyle       =   0  'Transparent
         Caption         =   "Photo Viewer"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Index           =   0
         Left            =   1290
         TabIndex        =   12
         Top             =   1185
         Width           =   1080
      End
      Begin VB.Shape shpTaskTileProjectName 
         BackColor       =   &H00AAAAAA&
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         Height          =   360
         Index           =   0
         Left            =   1170
         Shape           =   4  'Rounded Rectangle
         Top             =   1140
         Width           =   1320
      End
      Begin VB.Label lblTaskTileName 
         BackStyle       =   0  'Transparent
         Caption         =   "Team Meetup is already here"
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
         Left            =   1155
         TabIndex        =   11
         Top             =   195
         Width           =   3915
      End
      Begin VB.Label lblTaskTileTimeDesc 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "day(s) left"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Index           =   0
         Left            =   90
         TabIndex        =   10
         Top             =   1155
         Width           =   990
      End
      Begin VB.Label lblTaskTileTime 
         Alignment       =   2  'Center
         BackColor       =   &H00EAF3EC&
         BackStyle       =   0  'Transparent
         Caption         =   "2"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   48
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   990
         Index           =   0
         Left            =   150
         TabIndex        =   9
         Top             =   120
         Width           =   795
      End
   End
   Begin VB.PictureBox pbxPrjTile 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1410
      Index           =   0
      Left            =   210
      ScaleHeight     =   1410
      ScaleWidth      =   8520
      TabIndex        =   0
      Top             =   2220
      Visible         =   0   'False
      Width           =   8520
      Begin VB.Image imgPrjStar5 
         Height          =   345
         Index           =   0
         Left            =   1500
         Picture         =   "frmOverView.frx":0000
         Stretch         =   -1  'True
         Top             =   915
         Width           =   300
      End
      Begin VB.Image imgPrjStar4 
         Height          =   345
         Index           =   0
         Left            =   1170
         Picture         =   "frmOverView.frx":0452
         Stretch         =   -1  'True
         Top             =   915
         Width           =   300
      End
      Begin VB.Image imgPrjStar3 
         Height          =   345
         Index           =   0
         Left            =   870
         Picture         =   "frmOverView.frx":08A4
         Stretch         =   -1  'True
         Top             =   915
         Width           =   270
      End
      Begin VB.Image imgPrjStar2 
         Height          =   345
         Index           =   0
         Left            =   570
         Picture         =   "frmOverView.frx":0E06
         Stretch         =   -1  'True
         Top             =   915
         Width           =   270
      End
      Begin VB.Image ImgPrjStar1 
         Height          =   345
         Index           =   0
         Left            =   255
         Picture         =   "frmOverView.frx":1368
         Stretch         =   -1  'True
         Top             =   915
         Width           =   270
      End
      Begin VB.Label lblPrjProgressLbl 
         BackStyle       =   0  'Transparent
         Caption         =   "progress"
         BeginProperty Font 
            Name            =   "Agency FB"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   0
         Left            =   5730
         TabIndex        =   5
         Top             =   285
         Width           =   735
      End
      Begin VB.Label lblPrjProgress 
         BackStyle       =   0  'Transparent
         Caption         =   "N/A"
         BeginProperty Font 
            Name            =   "Agency FB"
            Size            =   24
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   615
         Index           =   0
         Left            =   5685
         TabIndex        =   4
         Top             =   600
         Width           =   915
      End
      Begin VB.Label lblPrjDesc 
         BackStyle       =   0  'Transparent
         Caption         =   "Simple photo viewer application"
         BeginProperty Font 
            Name            =   "Arial Unicode MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   0
         Left            =   225
         TabIndex        =   3
         Top             =   525
         Width           =   5055
      End
      Begin VB.Label lblPrjName 
         BackStyle       =   0  'Transparent
         Caption         =   "Photo Viewer"
         BeginProperty Font 
            Name            =   "Arial Rounded MT Bold"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   0
         Left            =   225
         TabIndex        =   2
         Top             =   255
         Width           =   4410
      End
      Begin VB.Label lblPrjTagName 
         BackStyle       =   0  'Transparent
         Caption         =   "C#"
         BeginProperty Font 
            Name            =   "Agency FB"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   0
         Left            =   1965
         TabIndex        =   1
         Top             =   945
         Width           =   255
      End
      Begin VB.Shape shpPrjTagName 
         BackColor       =   &H00AAAAAA&
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         Height          =   375
         Index           =   0
         Left            =   1860
         Shape           =   4  'Rounded Rectangle
         Top             =   900
         Width           =   495
      End
   End
   Begin VB.Label lblTaskHeader 
      BackStyle       =   0  'Transparent
      Caption         =   "UPCOMING TASKS"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   660
      Left            =   9150
      TabIndex        =   8
      Top             =   1485
      Width           =   6375
   End
   Begin VB.Image imgTaskMore 
      Height          =   525
      Left            =   15840
      Picture         =   "frmOverView.frx":18CA
      Stretch         =   -1  'True
      Top             =   1335
      Width           =   525
   End
   Begin VB.Image imgPrjMore 
      Height          =   525
      Left            =   7620
      Picture         =   "frmOverView.frx":49D5
      Stretch         =   -1  'True
      Top             =   1395
      Width           =   525
   End
   Begin VB.Label lblPrjHeader 
      BackStyle       =   0  'Transparent
      Caption         =   "ACTIVE PROJECTS"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   660
      Left            =   855
      TabIndex        =   6
      Top             =   1470
      Width           =   6375
   End
   Begin ComctlLib.ImageList imgStars 
      Left            =   135
      Top             =   195
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   48
      ImageHeight     =   48
      MaskColor       =   12632256
      UseMaskColor    =   0   'False
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   2
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmOverView.frx":7AE0
            Key             =   "off"
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmOverView.frx":9632
            Key             =   "on"
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmOverView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim prjManager As ProjectManager
Dim tmpProject As Project
Dim tskManager As TaskManager
Dim tmpTask As Task

Private Sub Form_Load()
    adjustContainers
    initProjectsSection
    initTasksSection
    pbxPrjTile(0).Left = -(pbxPrjTile(0).width)
    pbxTaskTile(0).Left = -(pbxTaskTile(0).width)
    
End Sub
Private Sub Form_Resize()
    Set prjManager = Nothing
    adjustContainers
    initProjectsSection
    initTasksSection
    pbxPrjTile(0).Left = -(pbxPrjTile(0).width)
    pbxTaskTile(0).Left = -(pbxTaskTile(0).width)
    
End Sub




Private Sub pbxPrjTile_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call changePrjTileStyle(Index, &HFFDDBB, False)
    
End Sub

Private Sub pbxPrjTile_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call changePrjTileStyle(Index, vbWhite, True)
End Sub

Private Sub pbxPrjTile_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    SetCursor LoadCursor(0, IDC_HAND)
End Sub

















'Functions

Private Sub adjustContainers()
    lblPrjHeader.Left = 560
    lblPrjHeader.width = pbxPrjTile(0).width - imgPrjMore.width - 365
    imgPrjMore.Left = lblPrjHeader.Left + lblPrjHeader.width + 365
    
    lblTaskHeader.Left = Me.width - pbxTaskTile(0).width - 560
    lblTaskHeader.width = pbxTaskTile(0).width - imgTaskMore.width - 365
    imgTaskMore.Left = lblTaskHeader.Left + lblTaskHeader.width + 365
    
End Sub

Private Sub initProjectsSection()

        Set prjManager = New ProjectManager
        Call prjManager.loadProjectsFromDB(" state = 'ACTIVE' ORDER BY last_update DESC ")
        prjManager.IsDebug = True
        If prjManager.IsLoaded Then
            Dim topThreeOrDefault As Integer
            topThreeOrDefault = prjManager.Count
            If prjManager.Count > 3 Then
                topThreeOrDefault = 3
            End If
            
            For i = 1 To topThreeOrDefault
                Call addProjectTile(i, prjManager.Projects(i).ProjectName, prjManager.Projects(i).Description, "0.0%", 0, prjManager.Projects(i).Tags)
                
            Next i
        End If
End Sub


Private Sub initTasksSection()

        Dim topThreeOrDefault As Integer
        Dim dtDiff As Variant
        Set tskManager = New TaskManager
        Call tskManager.loadTasksFromDB(" status = 'PENDING' and convert(date,CURRENT_TIMESTAMP) = convert(date,end_date) ORDER BY end_date ASC ")
        tskManager.IsDebug = True
        If tskManager.IsLoaded Then
            topThreeOrDefault = tskManager.Count
            If tskManager.Count > 3 Then
                topThreeOrDefault = 3
            End If
            For i = 1 To topThreeOrDefault
                dtDiff = DateDiff("d", DateTime.Date, tskManager.Tasks(i).EndDate, vbUseSystemDayOfWeek, vbUseSystem)
                'If dtDiff >= 0 Then
                    Call addTaskTile(i, tskManager.Tasks(i).TaskName, tskManager.Tasks(i).Description, dtDiff, "Empty")
                'End If
            Next i
            lblTaskHeader.Caption = "TODAY'S TASKS"
        End If
        If tskManager.Count < 1 Then
            Set tskManager = New TaskManager
            Call tskManager.loadTasksFromDB(" status = 'PENDING' and convert(date,CURRENT_TIMESTAMP) < convert(date,end_date) ORDER BY end_date ASC ")
            tskManager.IsDebug = True
            If tskManager.IsLoaded Then
                topThreeOrDefault = tskManager.Count
                If tskManager.Count > 3 Then
                    topThreeOrDefault = 3
                End If
                For i = 1 To topThreeOrDefault
                    dtDiff = DateDiff("d", DateTime.Date, tskManager.Tasks(i).EndDate, vbUseSystemDayOfWeek, vbUseSystem)
                    'If dtDiff >= 0 Then
                        Call addTaskTile(i, tskManager.Tasks(i).TaskName, tskManager.Tasks(i).Description, dtDiff, "Empty")
                    'End If
                Next i
            End If
            
            lblTaskHeader.Caption = "UPCOMING TASKS"
            
        End If
        
End Sub

Private Sub addProjectTile(ByVal Index As Integer, name As String, desc As String, progress As String, starCount As Integer, strTag As String, Optional tag As String = "")
    CreateProjectTile
    Call SetProjectTileInfo(Index, name, desc, progress, starCount, strTag, tag)
    
    pbxPrjTile(Index).Left = lblPrjHeader.Left
    If Index > 1 Then
        pbxPrjTile(Index).Top = pbxPrjTile(Index - 1).Top + pbxPrjTile(Index - 1).height + 200
    Else
        pbxPrjTile(Index).Top = lblPrjHeader.Top + lblPrjHeader.height + 200
        
    End If
    
    
End Sub

Private Function CreateProjectTile() As Integer
    Dim nextIndex As Integer
    nextIndex = pbxPrjTile.Count
    
    'load the contents
    Load pbxPrjTile(nextIndex)
    Load lblPrjName(nextIndex)
    Load lblPrjDesc(nextIndex)
    Load lblPrjProgressLbl(nextIndex)
    Load lblPrjProgress(nextIndex)
    Load lblPrjTagName(nextIndex)
    Load shpPrjTagName(nextIndex)
    Load ImgPrjStar1(nextIndex)
    Load imgPrjStar2(nextIndex)
    Load imgPrjStar3(nextIndex)
    Load imgPrjStar4(nextIndex)
    Load imgPrjStar5(nextIndex)
    
    Set lblPrjName(lblPrjName.Count - 1).Container = pbxPrjTile(nextIndex)
    Set lblPrjDesc(lblPrjDesc.Count - 1).Container = pbxPrjTile(nextIndex)
    Set lblPrjProgress(lblPrjProgress.Count - 1).Container = pbxPrjTile(nextIndex)
    Set lblPrjProgressLbl(lblPrjProgressLbl.Count - 1).Container = pbxPrjTile(nextIndex)
    Set lblPrjTagName(lblPrjTagName.Count - 1).Container = pbxPrjTile(nextIndex)
    Set shpPrjTagName(shpPrjTagName.Count - 1).Container = pbxPrjTile(nextIndex)
    Set ImgPrjStar1(ImgPrjStar1.Count - 1).Container = pbxPrjTile(nextIndex)
    Set imgPrjStar2(imgPrjStar2.Count - 1).Container = pbxPrjTile(nextIndex)
    Set imgPrjStar3(imgPrjStar3.Count - 1).Container = pbxPrjTile(nextIndex)
    Set imgPrjStar4(imgPrjStar4.Count - 1).Container = pbxPrjTile(nextIndex)
    Set imgPrjStar5(imgPrjStar5.Count - 1).Container = pbxPrjTile(nextIndex)
    
    
    
    pbxPrjTile(pbxPrjTile.Count - 1).Visible = True
    lblPrjName(lblPrjName.Count - 1).Visible = True
    lblPrjDesc(lblPrjDesc.Count - 1).Visible = True
    lblPrjProgress(lblPrjProgress.Count - 1).Visible = True
    lblPrjProgressLbl(lblPrjProgressLbl.Count - 1).Visible = True
    lblPrjTagName(lblPrjTagName.Count - 1).Visible = True
    shpPrjTagName(shpPrjTagName.Count - 1).Visible = True
    ImgPrjStar1(ImgPrjStar1.Count - 1).Visible = True
    imgPrjStar2(imgPrjStar2.Count - 1).Visible = True
    imgPrjStar3(imgPrjStar3.Count - 1).Visible = True
    imgPrjStar4(imgPrjStar4.Count - 1).Visible = True
    imgPrjStar5(imgPrjStar5.Count - 1).Visible = True
    
    CreateProjectTile = pbxPrjTile.Count - 1

End Function



Private Sub SetProjectTileInfo(Index As Integer, name As String, desc As String, progress As String, starCount As Integer, strTag As String, Optional tag As String = "")
    
    lblPrjName(Index).Caption = name
    lblPrjDesc(Index).Caption = desc
    lblPrjProgress(Index).Caption = progress
    lblPrjProgress(Index).width = Len(Replace(progress, " ", "")) * 200
    lblPrjProgress(Index).Left = pbxPrjTile(Index).width - lblPrjProgress(Index).width - 200
    lblPrjProgressLbl(Index).Left = lblPrjProgress(Index).Left
    lblPrjTagName(Index).Caption = strTag
    shpPrjTagName(Index).width = Len(Replace(strTag, " ", "")) * 100 + 100
    lblPrjTagName(Index).width = Len(Replace(strTag, " ", "")) * 100 + 100
    
    Call setStarCount(Index, starCount)
    
    If Len(tag) > 0 Then
        pbxPrjTile(Index).tag = tag
    End If

End Sub

Private Sub setStarCount(Index As Integer, starCount As Integer)

    ImgPrjStar1(Index).Picture = imgStars.ListImages("off").Picture
    imgPrjStar2(Index).Picture = imgStars.ListImages("off").Picture
    imgPrjStar3(Index).Picture = imgStars.ListImages("off").Picture
    imgPrjStar4(Index).Picture = imgStars.ListImages("off").Picture
    imgPrjStar5(Index).Picture = imgStars.ListImages("off").Picture
    
    If starCount >= 1 Then
        ImgPrjStar1(Index).Picture = imgStars.ListImages("on").Picture
    End If
    If starCount >= 2 Then
        imgPrjStar2(Index).Picture = imgStars.ListImages("on").Picture
    End If
    If starCount >= 3 Then
        imgPrjStar3(Index).Picture = imgStars.ListImages("on").Picture
    End If
    If starCount >= 4 Then
        imgPrjStar4(Index).Picture = imgStars.ListImages("on").Picture
    End If
    If starCount >= 5 Then
        imgPrjStar5(Index).Picture = imgStars.ListImages("on").Picture
    End If
    
    
End Sub

Private Sub changePrjTileStyle(Index As Integer, tBackColor As OLE_COLOR, hideStars As Boolean)
    pbxPrjTile(Index).BackColor = tBackColor
    
    ImgPrjStar1(Index).Visible = hideStars
    imgPrjStar2(Index).Visible = hideStars
    imgPrjStar3(Index).Visible = hideStars
    imgPrjStar4(Index).Visible = hideStars
    imgPrjStar5(Index).Visible = hideStars
End Sub
Private Sub changeTaskTileStyle(Index As Integer, tBackColor As OLE_COLOR)
    pbxTaskTile(Index).BackColor = tBackColor
End Sub



Private Sub addTaskTile(ByVal Index As Integer, ByVal name As String, ByVal desc As String, ByVal days As String, ByVal prjName As String, Optional tag As String = "")
    CreateTaskTile
    Call SetTaskTileInfo(Index, name, days, desc, prjName, tag)
    
    pbxTaskTile(Index).Left = lblTaskHeader.Left
    If Index > 1 Then
        pbxTaskTile(Index).Top = pbxTaskTile(Index - 1).Top + pbxTaskTile(Index - 1).height + 200
    Else
        pbxTaskTile(Index).Top = lblTaskHeader.Top + lblTaskHeader.height + 200
        
    End If
End Sub

Private Function CreateTaskTile() As Integer
    Dim nextIndex As Integer
    nextIndex = pbxTaskTile.Count
    
    'load the contents
    Load pbxTaskTile(nextIndex)
    Load lblTaskTileName(nextIndex)
    Load lblTaskTileDesc(nextIndex)
    Load lblTaskTileTime(nextIndex)
    Load shpTaskTileProjectName(nextIndex)
    Load lblTaskTileProjectName(nextIndex)
    Load lblTaskTileTimeDesc(nextIndex)
    
    Set lblTaskTileName(lblTaskTileName.Count - 1).Container = pbxTaskTile(nextIndex)
    Set lblTaskTileDesc(lblTaskTileDesc.Count - 1).Container = pbxTaskTile(nextIndex)
    Set lblTaskTileProjectName(lblTaskTileProjectName.Count - 1).Container = pbxTaskTile(nextIndex)
    Set shpTaskTileProjectName(shpTaskTileProjectName.Count - 1).Container = pbxTaskTile(nextIndex)
    Set lblTaskTileTime(lblTaskTileTime.Count - 1).Container = pbxTaskTile(nextIndex)
    Set lblTaskTileTimeDesc(lblTaskTileTimeDesc.Count - 1).Container = pbxTaskTile(nextIndex)
   
    
    pbxTaskTile(pbxTaskTile.Count - 1).Visible = True
    lblTaskTileName(lblTaskTileName.Count - 1).Visible = True
    lblTaskTileDesc(lblTaskTileDesc.Count - 1).Visible = True
    lblTaskTileProjectName(lblTaskTileProjectName.Count - 1).Visible = True
    lblTaskTileTime(lblTaskTileTime.Count - 1).Visible = True
    lblTaskTileTimeDesc(lblTaskTileTimeDesc.Count - 1).Visible = True
    shpTaskTileProjectName(shpTaskTileProjectName.Count - 1).Visible = True
    
    CreateTaskTile = pbxTaskTile.Count - 1

End Function



Private Sub SetTaskTileInfo(Index As Integer, name As String, days As String, desc As String, prjName As String, Optional tag As String = "")
    
    If days = "0" Then
        days = ""
        lblTaskTileTimeDesc(Index).Visible = False
        lblTaskTileName(Index).Top = 160
        lblTaskTileTime(Index).width = 0
        lblTaskTileDesc(Index).Top = lblTaskTileName(Index).Top + lblTaskTileName(Index).height
        
        shpTaskTileProjectName(Index).Top = lblTaskTileDesc(Index).Top + lblTaskTileDesc(Index).height
        lblTaskTileProjectName(Index).Top = lblTaskTileDesc(Index).Top + lblTaskTileDesc(Index).height + 30
        
        pbxTaskTile(Index).height = shpTaskTileProjectName(Index).Top + shpTaskTileProjectName(Index).height + 150
        
    Else
        lblTaskTileTime(Index).width = (Len(Replace(days, " ", "")) + 2) * 300
        lblTaskTileTimeDesc(Index).width = lblTaskTileTime(Index).width
    
        
    End If
    lblTaskTileTimeDesc(Index).Top = lblTaskTileTime(Index).height + lblTaskTileTime(Index).Top
    
    lblTaskTileProjectName(Index).Left = lblTaskTileTime(Index).width + lblTaskTileTime(Index).Left + 100
    shpTaskTileProjectName(Index).Left = lblTaskTileTime(Index).width + lblTaskTileTime(Index).Left
    lblTaskTileName(Index).Left = lblTaskTileTime(Index).width + lblTaskTileTime(Index).Left
    lblTaskTileDesc(Index).Left = lblTaskTileTime(Index).width + lblTaskTileTime(Index).Left
    
    
    lblTaskTileName(Index).Caption = name
    lblTaskTileTime(Index).Caption = days
    lblTaskTileDesc(Index).Caption = desc

    lblTaskTileProjectName(Index).Caption = prjName
    shpTaskTileProjectName(Index).width = Len(Replace(prjName, " ", "")) * 100 + 300
    lblTaskTileProjectName(Index).width = Len(Replace(prjName, " ", "")) * 100 + 100
    
    If Len(tag) > 0 Then
        pbxTaskTile(Index).tag = tag
    End If

End Sub


