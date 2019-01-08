VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmSnippetView 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   0  'None
   Caption         =   "Snippet"
   ClientHeight    =   6165
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11955
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6165
   ScaleWidth      =   11955
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox pbxInfoLoc 
      Appearance      =   0  'Flat
      BackColor       =   &H00FCFCFC&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   870
      Left            =   15
      ScaleHeight     =   870
      ScaleWidth      =   11820
      TabIndex        =   15
      Top             =   1185
      Visible         =   0   'False
      Width           =   11820
      Begin VB.Label lblInfoLang 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "TYPE"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   345
         Left            =   10725
         TabIndex        =   18
         Top             =   345
         Width           =   870
      End
      Begin VB.Label lblInfoDesc 
         BackStyle       =   0  'Transparent
         Caption         =   "Description about the tile"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   345
         TabIndex        =   17
         Top             =   540
         Width           =   8100
      End
      Begin VB.Label lblInfoHead 
         BackStyle       =   0  'Transparent
         Caption         =   "Snippet Name"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   315
         TabIndex        =   16
         Top             =   150
         Width           =   5940
      End
   End
   Begin VB.PictureBox pbxOption 
      BackColor       =   &H00EFEFEF&
      BorderStyle     =   0  'None
      Height          =   1215
      Index           =   3
      Left            =   10650
      ScaleHeight     =   1215
      ScaleWidth      =   1230
      TabIndex        =   12
      Top             =   -90
      Width           =   1230
      Begin VB.Label lblOptionIcon 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "X"
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
         Index           =   3
         Left            =   360
         TabIndex        =   14
         Top             =   120
         Width           =   495
      End
      Begin VB.Label lblOption 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Close View"
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
         Index           =   3
         Left            =   0
         TabIndex        =   13
         Top             =   840
         Width           =   1215
      End
      Begin VB.Shape shpOptionIcon 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         FillColor       =   &H008080FF&
         Height          =   735
         Index           =   3
         Left            =   150
         Shape           =   3  'Circle
         Top             =   15
         Width           =   975
      End
   End
   Begin VB.PictureBox pbxHead 
      Appearance      =   0  'Flat
      BackColor       =   &H00F7F7F7&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1110
      Left            =   -45
      ScaleHeight     =   1110
      ScaleWidth      =   11955
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   11955
      Begin VB.PictureBox pbxOption 
         BackColor       =   &H00EFEFEF&
         BorderStyle     =   0  'None
         Height          =   1215
         Index           =   2
         Left            =   2580
         ScaleHeight     =   1215
         ScaleWidth      =   1230
         TabIndex        =   9
         Top             =   0
         Width           =   1230
         Begin VB.Label lblOption 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Delete Snippet"
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
            Index           =   2
            Left            =   0
            TabIndex        =   11
            Top             =   840
            Width           =   1215
         End
         Begin VB.Label lblOptionIcon 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "D"
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
            Index           =   2
            Left            =   360
            TabIndex        =   10
            Top             =   120
            Width           =   495
         End
         Begin VB.Shape shpOptionIcon 
            BackColor       =   &H00E0E0E0&
            BackStyle       =   1  'Opaque
            BorderStyle     =   0  'Transparent
            FillColor       =   &H008080FF&
            Height          =   735
            Index           =   2
            Left            =   120
            Shape           =   3  'Circle
            Top             =   0
            Width           =   975
         End
      End
      Begin VB.PictureBox pbxOption 
         BackColor       =   &H00EFEFEF&
         BorderStyle     =   0  'None
         Height          =   1215
         Index           =   1
         Left            =   1290
         ScaleHeight     =   1215
         ScaleWidth      =   1230
         TabIndex        =   6
         Top             =   15
         Width           =   1230
         Begin VB.Label lblOptionIcon 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "E"
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
            Index           =   1
            Left            =   360
            TabIndex        =   8
            Top             =   120
            Width           =   495
         End
         Begin VB.Label lblOption 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Export"
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
            Index           =   1
            Left            =   0
            TabIndex        =   7
            Top             =   840
            Width           =   1215
         End
         Begin VB.Shape shpOptionIcon 
            BackColor       =   &H00E0E0E0&
            BackStyle       =   1  'Opaque
            BorderStyle     =   0  'Transparent
            FillColor       =   &H008080FF&
            Height          =   735
            Index           =   1
            Left            =   120
            Shape           =   3  'Circle
            Top             =   0
            Width           =   975
         End
      End
      Begin VB.PictureBox pbxOption 
         BackColor       =   &H00EFEFEF&
         BorderStyle     =   0  'None
         Height          =   1215
         Index           =   0
         Left            =   0
         ScaleHeight     =   1215
         ScaleWidth      =   1230
         TabIndex        =   3
         Top             =   0
         Width           =   1230
         Begin VB.Label lblOptionIcon 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "S"
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
            Left            =   360
            TabIndex        =   5
            Top             =   120
            Width           =   495
         End
         Begin VB.Label lblOption 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Save Snippet"
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
            Left            =   0
            TabIndex        =   4
            Top             =   840
            Width           =   1215
         End
         Begin VB.Shape shpOptionIcon 
            BackColor       =   &H00E0E0E0&
            BackStyle       =   1  'Opaque
            BorderStyle     =   0  'Transparent
            FillColor       =   &H008080FF&
            Height          =   735
            Index           =   0
            Left            =   120
            Shape           =   3  'Circle
            Top             =   0
            Width           =   975
         End
      End
   End
   Begin RichTextLib.RichTextBox rtbSnippetView 
      Height          =   4365
      Left            =   600
      TabIndex        =   1
      Top             =   2460
      Visible         =   0   'False
      Width           =   11595
      _ExtentX        =   20452
      _ExtentY        =   7699
      _Version        =   393217
      BorderStyle     =   0
      ScrollBars      =   3
      Appearance      =   0
      TextRTF         =   $"frmSnippetView.frx":0000
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Consolas"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label lblLineNos 
      BackColor       =   &H00FCFCFC&
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4410
      Left            =   -45
      TabIndex        =   2
      Top             =   1680
      Width           =   435
   End
End
Attribute VB_Name = "frmSnippetView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public SnippetId As Integer


Private fso As New FileSystemObject
Private snippetObj As Snippet

Private fileName As String

Private Sub Form_Load()
    Call alignContainers
    Call SetupSnippet
End Sub
Private Sub Form_Resize()
    Call alignContainers
End Sub


Private Sub pbxOption_Click(Index As Integer)
    ActionHeader (Index)
    
End Sub

Private Sub lblOptionIcon_Click(Index As Integer)
    ActionHeader (Index)
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

Private Sub ActionHeader(Index As Integer)
    Select Case (Index)
        Case 0:
            Dim fileNo As Integer
            fileNo = FreeFile
            Open fileName For Output As #fileNo
            Print #fileNo, rtbSnippetView.Text
            Close #fileNo
    End Select
    MsgBox Index
End Sub

Private Sub SetupSnippet()
    If SnippetId > 0 Then
        Set snippetObj = New Snippet
        snippetObj.LoadSingleton (SnippetId)
        snippetObj.debugSnippet
        If snippetObj.IsLoaded Then
            
            lblInfoHead.Caption = snippetObj.SnippetName
            lblInfoDesc.Caption = snippetObj.Description
            lblInfoLang.Caption = UCase(snippetObj.snippetType)
            If fso.FileExists(snippetObj.Location) Then
                rtbSnippetView.fileName = snippetObj.Location
            End If
            fileName = snippetObj.Location
            
        End If
        
    Else
        MsgBox "Message Object Error", vbCritical
        
    End If
End Sub

Private Sub alignContainers()
    
    'Header PictureBox
    pbxHead.Top = 0
    pbxHead.height = pbxOption(0).height + 100
    For i = 0 To pbxOption.Count - 1
        pbxOption(i).Top = 50
        shpOptionIcon(i).Top = 50
        lblOptionIcon(i).Top = 150
    Next i
    'LineNumberLabel
    lblLineNos.Left = 0
    lblLineNos.height = Me.height - pbxHead.height
    lblLineNos.Top = pbxHead.height + pbxHead.Top
    lblLineNos.width = 360
    
    'Header
    pbxHead.width = Me.width
    pbxOption(pbxOption.Count - 1).Left = pbxHead.width - pbxOption(pbxOption.Count - 1).width - 100
    
    'info tile
    pbxInfoLoc.Top = pbxHead.height + pbxHead.Top
    pbxInfoLoc.Left = 0
    pbxInfoLoc.width = Me.width
    
    lblInfoLang.Left = pbxInfoLoc.width - lblInfoLang.width
    
    'RichTextBox
    rtbSnippetView.Left = 360
    rtbSnippetView.width = Me.width - lblLineNos.width
    rtbSnippetView.Top = pbxInfoLoc.height + pbxInfoLoc.Top
    rtbSnippetView.height = Me.height - pbxHead.height - pbxInfoLoc.height
    
    rtbSnippetView.Visible = True
    pbxHead.Visible = True
    pbxInfoLoc.Visible = True
End Sub

Private Sub rtbSnippetView_Change()
    'disable this function for now
    Exit Sub
    
    Dim rtbText As String
    Dim sliceArray() As String
    rtbText = rtbSnippetView.Text
    sliceArray = Split(rtbText, vbLf)
    lblLineNos.Caption = ""
    For i = 1 To UBound(sliceArray)
        lblLineNos.Caption = lblLineNos.Caption & i & " " & vbLf
    Next i


End Sub
