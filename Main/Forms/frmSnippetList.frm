VERSION 5.00
Begin VB.Form frmSnippetList 
   BackColor       =   &H00EFEFEF&
   BorderStyle     =   0  'None
   Caption         =   "Snippets"
   ClientHeight    =   9300
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   18870
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9300
   ScaleWidth      =   18870
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox pbxContent 
      BorderStyle     =   0  'None
      Height          =   8070
      Left            =   525
      ScaleHeight     =   8070
      ScaleWidth      =   11205
      TabIndex        =   3
      Top             =   1335
      Width           =   11205
      Begin VB.VScrollBar vsrContainer 
         Height          =   3135
         Left            =   10500
         SmallChange     =   200
         TabIndex        =   10
         Top             =   105
         Visible         =   0   'False
         Width           =   225
      End
      Begin VB.PictureBox pbxContainer 
         Appearance      =   0  'Flat
         BackColor       =   &H00EFEFEF&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   4350
         Left            =   525
         ScaleHeight     =   4350
         ScaleWidth      =   11655
         TabIndex        =   4
         Top             =   1110
         Width           =   11655
         Begin VB.PictureBox pbxSnippetTile 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   1320
            Index           =   0
            Left            =   450
            ScaleHeight     =   1320
            ScaleWidth      =   7965
            TabIndex        =   5
            Top             =   2505
            Visible         =   0   'False
            Width           =   7965
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
               Left            =   6705
               TabIndex        =   8
               Top             =   180
               Width           =   1095
            End
            Begin VB.Label lblSnippetName 
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
               TabIndex        =   7
               Top             =   270
               Width           =   4530
            End
            Begin VB.Label lblSnippetDesc 
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
            TabIndex        =   9
            Top             =   1845
            Width           =   6375
         End
      End
   End
   Begin VB.PictureBox pbxHead 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      CausesValidation=   0   'False
      ForeColor       =   &H80000008&
      Height          =   1455
      Left            =   345
      ScaleHeight     =   1425
      ScaleWidth      =   18840
      TabIndex        =   0
      Top             =   -105
      Width           =   18870
      Begin VB.PictureBox pbxOption 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   1215
         Index           =   0
         Left            =   17295
         ScaleHeight     =   1215
         ScaleWidth      =   1230
         TabIndex        =   11
         Top             =   165
         Width           =   1230
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
            TabIndex        =   13
            Top             =   870
            Width           =   1215
         End
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
            TabIndex        =   12
            Top             =   195
            Width           =   495
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
      Begin VB.Label lblHeadDesc 
         BackStyle       =   0  'Transparent
         Caption         =   "Snippets are short text files to store small things like code, note etc.. So that you can View them again "
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
         TabIndex        =   2
         Top             =   735
         Width           =   7830
      End
      Begin VB.Label lblSnippetHeader 
         BackStyle       =   0  'Transparent
         Caption         =   "Snippets"
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
         TabIndex        =   1
         Top             =   330
         Width           =   6375
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
Dim loaded As Boolean
Private Sub Form_Load()
    adjustContainers
    'initSnippetList
    pbxSnippetTile(0).Left = -(pbxSnippetTile(0).width)
    
End Sub
Private Sub Form_Resize()
    Set sniManager = Nothing
    If Me.WindowState = 2 Then
        adjustContainers
        alignControls
        initSnippetList
        pbxSnippetTile(0).Left = -(pbxSnippetTile(0).width)
    End If
    
End Sub






Private Sub lblSnippetView_Click(Index As Integer)
'Main Navigation
    frmMDIMain.hideAllWindows
    frmSnippetView.SnippetId = Val(pbxSnippetTile(Index).tag)
    
    frmSnippetView.Show
    frmSnippetView.WindowState = 2
End Sub

Private Sub lblSnippetView_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call changeSnippetStyle(Index, &HFFDDBB)
    
End Sub

Private Sub lblSnippetView_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call changeSnippetStyle(Index, vbWhite)
End Sub

Private Sub lblSnippetView_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    SetCursor LoadCursor(0, IDC_HAND)
End Sub

Private Sub vsrContainer_Change()
    Dim changeVal As Long
    changeVal = -vsrContainer.value
    
    If moveTop < pbxContainer.height Then
        pbxContainer.Top = changeVal& * 10
    End If
End Sub



Private Sub changeSnippetStyle(Index As Integer, tBackColor As OLE_COLOR)
    pbxSnippetTile(Index).BackColor = tBackColor
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
    
    frmNewSnippet.Show
    SetTopMostWindow frmNewSnippet.hwnd, True
End Sub

Private Sub lblOptionIcon_Click(Index As Integer)
    
    frmNewSnippet.Show
    SetTopMostWindow frmNewSnippet.hwnd, True
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

Private Sub initSnippetList()

        Set sniManager = New Snippets
        Call sniManager.loadSnippetsFromDB(" 1=1 ORDER BY type ")
        Call alignControls
        If sniManager.IsLoaded Then
            Dim snippetType As String
            For i = 1 To sniManager.Count
                
                snippetType = UCase(sniManager.Snippets(i).snippetType)
                If i > 1 Then
                    If UCase(sniManager.Snippets(i - 1).snippetType) = UCase(sniManager.Snippets(i).snippetType) Then
                        snippetType = ""
                    End If
                End If
                Call addSnippetTile(i, sniManager.Snippets(i).SnippetName, sniManager.Snippets(i).Description, snippetType, sniManager.Snippets(i).SnippetId)
                
            Next i
        End If
End Sub

Private Sub addSnippetTile(ByVal Index As Integer, ByVal name As String, ByVal desc As String, snippetType As String, Optional tag As String = "")
    CreateSnippetTile
    Call SetSnippetTileInfo(Index, name, desc, tag)
    If Index > 1 Then
        Dim nextPos As Long
        nextPos = pbxSnippetTile(Index - 1).Top + pbxSnippetTile(Index - 1).height + 200
    
        If snippetType <> "" Then
            Call CreateTypeHeader(lblTypeHeader.Count, snippetType)
            Call AdjustTypeHeader(lblTypeHeader.Count - 1, nextPos)
            nextPos = nextPos + lblTypeHeader(lblTypeHeader.Count - 1).height + 200
        End If
    
        pbxSnippetTile(Index).Top = nextPos
    Else
        lblTypeHeader(0).Caption = snippetType
        lblTypeHeader(0).Visible = True
        Call AdjustTypeHeader(0, lblSnippetHeader.Top + lblSnippetHeader.height + 260)
        pbxSnippetTile(1).Top = lblTypeHeader(0).height + lblTypeHeader(0).Top
        
    End If
    Call AdjustSnippetTile(Index)
    Call AdjustContainer(Index)
    
End Sub
Private Sub CreateTypeHeader(Index As Integer, snippetType As String)
    Load lblTypeHeader(Index)
    lblTypeHeader(Index).Caption = snippetType
    lblTypeHeader(Index).Visible = True
    
            
End Sub

Private Function CreateSnippetTile() As Integer
    Dim nextIndex As Integer
    nextIndex = pbxSnippetTile.Count
    
    'load the contents
    Load pbxSnippetTile(nextIndex)
    Load lblSnippetName(nextIndex)
    Load lblSnippetDesc(nextIndex)
    Load lblSnippetView(nextIndex)
    
    Set lblSnippetName(lblSnippetName.Count - 1).Container = pbxSnippetTile(nextIndex)
    Set lblSnippetDesc(lblSnippetDesc.Count - 1).Container = pbxSnippetTile(nextIndex)
    Set lblSnippetView(lblSnippetView.Count - 1).Container = pbxSnippetTile(nextIndex)
    
    pbxSnippetTile(pbxSnippetTile.Count - 1).Visible = True
    lblSnippetName(lblSnippetName.Count - 1).Visible = True
    lblSnippetDesc(lblSnippetDesc.Count - 1).Visible = True
    lblSnippetView(lblSnippetView.Count - 1).Visible = True
    
    CreateSnippetTile = pbxSnippetTile.Count - 1

End Function

Private Sub AdjustTypeHeader(Index As Integer, nextPos As Long)
    lblTypeHeader(Index).Top = nextPos + 50
    lblTypeHeader(Index).Left = pbxContainer.width * 0.1
End Sub

Private Sub AdjustSnippetTile(Index As Integer)
    
    
    pbxSnippetTile(Index).Left = pbxContainer.width * 0.1
    pbxSnippetTile(Index).width = pbxContainer.width * 0.8
    
    lblSnippetView(Index).Left = pbxSnippetTile(Index).width - lblSnippetView(Index).width - 100
    
End Sub


Private Sub SetSnippetTileInfo(Index As Integer, name As String, desc As String, Optional tag As String = "")
    
    lblSnippetName(Index).Caption = name
    lblSnippetDesc(Index).Caption = desc
    
    If Len(tag) > 0 Then
        pbxSnippetTile(Index).tag = tag
    End If

End Sub


Private Sub AdjustContainer(Index As Integer)
    
    pbxContainer.height = pbxSnippetTile(Index).Top + pbxSnippetTile(Index).height + 500
    
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
    i = lblSnippetName.Count - 1
    While i > 1
        Unload lblSnippetName(i)
        i = i - 1
    Wend
    
    i = lblSnippetDesc.Count - 1
    While i > 1
        Unload lblSnippetDesc(i)
        i = i - 1
    Wend
    
    i = lblSnippetView.Count - 1
    While i > 1
        Unload lblSnippetView(i)
        i = i - 1
    Wend
    
    i = lblTypeHeader.Count - 1
    While i > 1
        Unload lblTypeHeader(i)
        i = i - 1
    Wend
    
    i = pbxSnippetTile.Count
    While i > 1
        Unload pbxSnippetTile(i)
        i = i - 1
    Wend
    
    
    
End Sub

Private Sub alignControls()
    AdjustContainer (pbxSnippetTile.Count - 1)
End Sub
