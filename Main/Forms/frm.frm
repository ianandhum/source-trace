VERSION 5.00
Begin VB.Form frmTaskView 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Tasks"
   ClientHeight    =   8805
   ClientLeft      =   5175
   ClientTop       =   2820
   ClientWidth     =   18705
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8805
   ScaleWidth      =   18705
   ShowInTaskbar   =   0   'False
   Tag             =   "FormHost"
   Begin VB.PictureBox pbxContent 
      Appearance      =   0  'Flat
      BackColor       =   &H00FCFCFC&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   6495
      Left            =   240
      ScaleHeight     =   6495
      ScaleWidth      =   18255
      TabIndex        =   1
      Top             =   1920
      Width           =   18255
      Begin VB.PictureBox pbxTaskTile 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         DragIcon        =   "frm.frx":0000
         DragMode        =   1  'Automatic
         ForeColor       =   &H80000008&
         Height          =   1215
         Index           =   0
         Left            =   13800
         ScaleHeight     =   1215
         ScaleWidth      =   3615
         TabIndex        =   4
         Top             =   120
         Width           =   3615
         Begin VB.Label lblTaskTileHeader 
            BackStyle       =   0  'Transparent
            Caption         =   "TaskTileHeadTemplate"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   0
            Left            =   120
            TabIndex        =   7
            Top             =   120
            Width           =   1935
         End
         Begin VB.Label lblTaskTileContent 
            BackStyle       =   0  'Transparent
            Caption         =   "TaskTileContentTemplate"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   0
            Left            =   120
            TabIndex        =   6
            Top             =   480
            Width           =   3375
         End
         Begin VB.Label lblTaskTileDate 
            Caption         =   "  TaskTileDate"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   5
            Top             =   840
            Width           =   975
         End
      End
      Begin VB.PictureBox pbxCard 
         Appearance      =   0  'Flat
         BackColor       =   &H00EEEEEE&
         BorderStyle     =   0  'None
         DragIcon        =   "frm.frx":C922
         DragMode        =   1  'Automatic
         ForeColor       =   &H80000008&
         Height          =   5175
         Index           =   0
         Left            =   720
         ScaleHeight     =   5175
         ScaleWidth      =   4215
         TabIndex        =   2
         Tag             =   "Card"
         Top             =   360
         Width           =   4215
         Begin VB.Label lblCardHeader 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "CardHeaderTemplate"
            BeginProperty Font 
               Name            =   "Consolas"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   0
            Left            =   0
            TabIndex        =   3
            Top             =   240
            Width           =   4215
         End
      End
   End
   Begin VB.PictureBox pbxHead 
      Appearance      =   0  'Flat
      BackColor       =   &H00EFEFEF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1575
      Left            =   240
      ScaleHeight     =   1575
      ScaleWidth      =   18225
      TabIndex        =   0
      Top             =   120
      Width           =   18225
   End
End
Attribute VB_Name = "frmTaskView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim preserveBackColor
    

Private Sub pbxCard_DragDrop(Index As Integer, Source As Control, X As Single, Y As Single)
    If Source.Tag = "Card" Then Exit Sub
    
    Set Source.Container = pbxCard(Index)
    Source.Left = 360
    Source.Top = Y - 100
    pbxCard(Index).BackColor = preserveBackColor
    
    
    
End Sub

Private Sub pbxCard_DragOver(Index As Integer, Source As Control, X As Single, Y As Single, State As Integer)
    
    For i = 0 To pbxCard.Count - 1
        If i <> Index Then
            pbxCard(i).BackColor = preserveBackColor
        End If
    Next i
    pbxCard(Index).BackColor = &H77FFDD
    
End Sub

Private Sub pbxContent_DragDrop(Source As Control, X As Single, Y As Single)
    
    If Source.Tag = "Card" Then
        Source.Left = X
    End If
        
End Sub
Private Sub Form_Load()
    alignContainers
    
    preserveBackColor = pbxCard(0).BackColor
    
    Dim lastIndex As Integer
    
    lastIndex = CreateTaskTile()
    
    Set pbxTaskTile(lastIndex).Container = pbxCard(0)
    
    pbxTaskTile(lastIndex).Top = 660
    pbxTaskTile(lastIndex).Left = 360
    SetTileInfo lastIndex, "This Head", "Content man issue", "Nov 19 2018"
    
    For i = 1 To 2
        lastIndex = CreateCard()
        
        Set pbxCard(lastIndex).Container = pbxContent
        
        pbxCard(lastIndex).Top = pbxCard(i - 1).Top
        pbxCard(lastIndex).Left = pbxCard(i - 1).Left + pbxCard(i - 1).width + 560
    Next i
End Sub
Private Sub Form_Resize()
    alignContainers
End Sub


'Functions

Private Sub alignContainers()
    pbxHead.Top = 100
    pbxHead.Left = 100
    pbxHead.width = Me.width - 200
    pbxHead.height = 1300
    pbxContent.Top = 100 + pbxHead.Top + pbxHead.height
    pbxContent.Left = 100
    pbxContent.width = Me.width - 200
    pbxContent.height = Me.height - pbxContent.Top - 100
    
End Sub
Private Function CreateCard() As Integer
    
    'load the contents
    Load pbxCard(pbxCard.Count)
    Load lblCardHeader(lblCardHeader.Count)
    
    Set lblCardHeader(lblCardHeader.Count - 1).Container = pbxCard(pbxCard.Count - 1)
    
    pbxCard(pbxCard.Count - 1).Visible = True
    lblCardHeader(lblCardHeader.Count - 1).Visible = True

    CreateCard = pbxCard.Count - 1
End Function

Private Function CreateTaskTile() As Integer
    
    'load the contents
    Load pbxTaskTile(pbxTaskTile.Count)
    Load lblTaskTileHeader(lblTaskTileHeader.Count)
    Load lblTaskTileContent(lblTaskTileContent.Count)
    Load lblTaskTileDate(lblTaskTileDate.Count)
    
    Set lblTaskTileHeader(lblTaskTileHeader.Count - 1).Container = pbxTaskTile(pbxTaskTile.Count - 1)
    Set lblTaskTileContent(lblTaskTileContent.Count - 1).Container = pbxTaskTile(pbxTaskTile.Count - 1)
    Set lblTaskTileDate(lblTaskTileDate.Count - 1).Container = pbxTaskTile(pbxTaskTile.Count - 1)
    
    pbxTaskTile(pbxTaskTile.Count - 1).Visible = True
    lblTaskTileHeader(lblTaskTileHeader.Count - 1).Visible = True
    lblTaskTileContent(lblTaskTileContent.Count - 1).Visible = True
    lblTaskTileDate(lblTaskTileDate.Count - 1).Visible = True
    
    CreateTaskTile = pbxTaskTile.Count - 1
End Function

Private Sub SetTileInfo(Index As Integer, header As String, content As String, dt As String)
    lblTaskTileHeader(Index).Caption = header
    lblTaskTileContent(Index).Caption = content
    lblTaskTileDate(Index).Caption = "  " & dt
End Sub

