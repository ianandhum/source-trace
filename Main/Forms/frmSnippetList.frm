VERSION 5.00
Begin VB.Form frmSnippetList 
   BackColor       =   &H00EFEFEF&
   BorderStyle     =   0  'None
   Caption         =   "Snippets"
   ClientHeight    =   6300
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12135
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6300
   ScaleWidth      =   12135
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox pbxSnippetTile 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1320
      Index           =   0
      Left            =   885
      ScaleHeight     =   1320
      ScaleWidth      =   8850
      TabIndex        =   0
      Top             =   2595
      Visible         =   0   'False
      Width           =   8850
      Begin VB.Label lblSnippetView 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "View Snippet"
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
         Left            =   7650
         TabIndex        =   4
         Top             =   120
         Width           =   1095
      End
      Begin VB.Label lblSnippetName 
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
         ForeColor       =   &H00606060&
         Height          =   435
         Index           =   0
         Left            =   180
         TabIndex        =   3
         Top             =   75
         Width           =   3915
      End
      Begin VB.Label lblSnippetType 
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
         Left            =   315
         TabIndex        =   2
         Top             =   900
         Width           =   1080
      End
      Begin VB.Label lblSnippetDesc 
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
         ForeColor       =   &H00606060&
         Height          =   285
         Index           =   0
         Left            =   180
         TabIndex        =   1
         Top             =   510
         Width           =   8460
      End
      Begin VB.Shape shpSnippetType 
         BackColor       =   &H00AAAAAA&
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         Height          =   360
         Index           =   0
         Left            =   195
         Shape           =   4  'Rounded Rectangle
         Top             =   855
         Width           =   1320
      End
   End
End
Attribute VB_Name = "frmSnippetList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim sniManager As Snippets
Dim tmpSnippet As Snippet

Private Sub Form_Load()
    adjustContainers
    initSnippetList
    pbxSnippetTile(0).Left = -(pbxSnippetTile(0).width)
    
End Sub
Private Sub Form_Resize()
    Set sniManager = Nothing
    adjustContainers
    initSnippetList
    pbxSnippetTile(0).Left = -(pbxSnippetTile(0).width)
    
End Sub




Private Sub pbxSnippetTile_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call changeSnippetStyle(Index, &HFFAA77)
    
End Sub

Private Sub pbxSnippetTile_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call changeSnippetStyle(Index, vbWhite)
End Sub

Private Sub pbxSnippetTile_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    SetCursor LoadCursor(0, IDC_HAND)
End Sub



Private Sub changeSnippetStyle(Index As Integer, tBackColor As OLE_COLOR)
    pbxSnippetTile(Index).BackColor = tBackColor
End Sub















'Functions

Private Sub adjustContainers()
    
End Sub

Private Sub initSnippetList()

        Set sniManager = New Snippets
        Call sniManager.loadSnippetsFromDB(" 1=1 ORDER BY snippet_id DESC ")
        sniManager.IsDebug = True
        If sniManager.IsLoaded Then
            Dim topThreeOrDefault As Integer
            topThreeOrDefault = sniManager.Count
            If sniManager.Count > 3 Then
                topThreeOrDefault = 3
            End If
            
            For i = 1 To topThreeOrDefault
                Call addSnippetTile(i, sniManager.Snippets(i).SnippetName, sniManager.Snippets(i).Description, sniManager.Snippets(i).SnippetType)
                
            Next i
        End If
End Sub

Private Sub addSnippetTile(ByVal Index As Integer, ByVal name As String, ByVal desc As String, ByVal typeName As String, Optional tag As String = "")
    CreateSnippetTile
    Call SetSnippetTileInfo(Index, name, desc, typeName, tag)
    
    'pbxSnippetTile(Index).Left = lblSnippetHeader.Left
    If Index > 1 Then
        pbxSnippetTile(Index).Top = pbxSnippetTile(Index - 1).Top + pbxSnippetTile(Index - 1).height + 200
    Else
        'pbxSnippetTile(Index).Top = lblSnippetHeader.Top + lblSnippetHeader.height + 200
        
    End If
End Sub

Private Function CreateSnippetTile() As Integer
    Dim nextIndex As Integer
    nextIndex = pbxSnippetTile.Count
    
    'load the contents
    Load pbxSnippetTile(nextIndex)
    Load lblSnippetName(nextIndex)
    Load lblSnippetDesc(nextIndex)
    Load lblSnippetType(nextIndex)
    Load shpSnippetType(nextIndex)
    Load lblSnippetView(nextIndex)
    
    Set lblSnippetName(lblSnippetName.Count - 1).Container = pbxSnippetTile(nextIndex)
    Set lblSnippetDesc(lblSnippetDesc.Count - 1).Container = pbxSnippetTile(nextIndex)
    Set lblSnippetType(lblSnippetType.Count - 1).Container = pbxSnippetTile(nextIndex)
    Set shpSnippetType(shpSnippetType.Count - 1).Container = pbxSnippetTile(nextIndex)
    Set lblSnippetView(lblSnippetView.Count - 1).Container = pbxSnippetTile(nextIndex)
    
    pbxSnippetTile(pbxSnippetTile.Count - 1).Visible = True
    lblSnippetName(lblSnippetName.Count - 1).Visible = True
    lblSnippetDesc(lblSnippetDesc.Count - 1).Visible = True
    lblSnippetType(lblSnippetType.Count - 1).Visible = True
    lblSnippetView(lblSnippetView.Count - 1).Visible = True
    shpSnippetType(shpSnippetType.Count - 1).Visible = True
    
    CreateSnippetTile = pbxSnippetTile.Count - 1

End Function



Private Sub SetSnippetTileInfo(Index As Integer, name As String, desc As String, typeName As String, Optional tag As String = "")
    
    Exit Sub
    lblSnippetName(Index).Caption = name
    lblSnippetDesc(Index).Caption = desc
    
    lblSnippetType(Index).Caption = typeName
    shpSnippetType(Index).width = Len(Replace(sniName, " ", "")) * 100 + 300
    lblSnippetType(Index).width = Len(Replace(sniName, " ", "")) * 100 + 100
    
    If Len(tag) > 0 Then
        pbxSnippetTile(Index).tag = tag
    End If

End Sub



