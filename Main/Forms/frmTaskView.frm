VERSION 5.00
Begin VB.Form frmTaskView 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Tasks"
   ClientHeight    =   11295
   ClientLeft      =   5175
   ClientTop       =   2820
   ClientWidth     =   18705
   BeginProperty Font 
      Name            =   "Arial Narrow"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   11295
   ScaleWidth      =   18705
   ShowInTaskbar   =   0   'False
   Tag             =   "FormHost"
   Begin VB.HScrollBar hscrContent 
      Height          =   255
      LargeChange     =   500
      Left            =   0
      SmallChange     =   100
      TabIndex        =   14
      Top             =   11040
      Visible         =   0   'False
      Width           =   4095
   End
   Begin VB.PictureBox pbxContent 
      Appearance      =   0  'Flat
      BackColor       =   &H00FCFCFC&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   9255
      Left            =   240
      ScaleHeight     =   9255
      ScaleWidth      =   18255
      TabIndex        =   1
      Top             =   1920
      Width           =   18255
      Begin VB.TextBox txtTempEdit 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   375
         Left            =   10680
         TabIndex        =   13
         Top             =   1080
         Visible         =   0   'False
         Width           =   2415
      End
      Begin VB.PictureBox pbxNewCard 
         Appearance      =   0  'Flat
         BackColor       =   &H00EEEEEE&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   5160
         ScaleHeight     =   615
         ScaleWidth      =   4215
         TabIndex        =   11
         Top             =   360
         Width           =   4215
         Begin VB.Label lblnewCardHead 
            Alignment       =   2  'Center
            Caption         =   " +  Add Another Card"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   375
            Left            =   240
            TabIndex        =   12
            Top             =   120
            Width           =   3735
         End
      End
      Begin VB.PictureBox pbxTaskTile 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   1215
         Index           =   0
         Left            =   13800
         ScaleHeight     =   1215
         ScaleWidth      =   3615
         TabIndex        =   4
         Tag             =   "Tile"
         Top             =   120
         Visible         =   0   'False
         Width           =   3615
         Begin VB.Label lblTileHandle 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Move"
            DragIcon        =   "frmTaskView.frx":0000
            DragMode        =   1  'Automatic
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Index           =   0
            Left            =   2775
            TabIndex        =   19
            Top             =   105
            Width           =   810
         End
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
            Width           =   2595
            WordWrap        =   -1  'True
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
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   2175
         Index           =   0
         Left            =   720
         ScaleHeight     =   2175
         ScaleWidth      =   4215
         TabIndex        =   2
         Tag             =   "Card"
         Top             =   360
         Visible         =   0   'False
         Width           =   4215
         Begin VB.Label lblCardAddNew 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "+ Add New Task"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF8080&
            Height          =   495
            Index           =   0
            Left            =   0
            TabIndex        =   8
            Top             =   1680
            Width           =   4215
         End
         Begin VB.Label lblCardHeader 
            BackStyle       =   0  'Transparent
            Caption         =   "CardHeaderTemplate"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   0
            Left            =   360
            TabIndex        =   3
            Top             =   120
            Width           =   2655
         End
      End
      Begin VB.Label lblDeleteNotice 
         BackStyle       =   0  'Transparent
         Caption         =   "Note: Task Can Be deleted By dropping it to the container itself"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   540
         Left            =   435
         TabIndex        =   18
         Top             =   8265
         Visible         =   0   'False
         Width           =   8220
      End
   End
   Begin VB.PictureBox pbxHead 
      Appearance      =   0  'Flat
      BackColor       =   &H00EFEFEF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   1575
      Left            =   240
      ScaleHeight     =   1575
      ScaleWidth      =   18225
      TabIndex        =   0
      Top             =   120
      Width           =   18225
      Begin VB.PictureBox pbxReset 
         BackColor       =   &H00EFEFEF&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1215
         Left            =   15360
         ScaleHeight     =   1215
         ScaleWidth      =   1230
         TabIndex        =   15
         Top             =   240
         Width           =   1230
         Begin VB.Label lblReset 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Reload Tasks"
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
            Left            =   0
            TabIndex        =   17
            Top             =   840
            Width           =   1215
         End
         Begin VB.Label lblResetIcon 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "R"
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
            Left            =   360
            TabIndex        =   16
            Top             =   120
            Width           =   495
         End
         Begin VB.Shape shpResetIcon 
            BackColor       =   &H00E0E0E0&
            BackStyle       =   1  'Opaque
            BorderStyle     =   0  'Transparent
            FillColor       =   &H008080FF&
            Height          =   735
            Left            =   120
            Shape           =   3  'Circle
            Top             =   0
            Width           =   975
         End
      End
      Begin VB.Label lblPrjDesc 
         Caption         =   "It is a simple application to view and orgainize  data "
         Height          =   495
         Left            =   240
         TabIndex        =   10
         Top             =   600
         Width           =   10815
      End
      Begin VB.Label lblPrjName 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "PhotoViewer"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   375
         Left            =   240
         TabIndex        =   9
         Top             =   240
         Width           =   8175
      End
   End
End
Attribute VB_Name = "frmTaskView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'default color of the card
Dim preserveBackColor

'Number of tasks in each card
Dim tileCounts() As Integer
Dim tCards As TaskCards
Dim cCardManager As Collection
'Status of txtTmpEdit
Dim TempEditing As Boolean
Private Sub Form_Load()
    alignContainers
    preserveBackColor = pbxCard(0).BackColor
    TempEditing = False
    initTaskView
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    ReDim tileCounts(1)
    Set cCardManager = Nothing
    Set tCards = Nothing
End Sub

Private Sub Form_Resize()
    alignContainers

End Sub


Private Sub lblCardAddNew_Click(Index As Integer)
    frmNewTask.CardIndex = Index
    frmNewTask.Show
End Sub

Public Sub addNewTile(Index As Integer, header As String, content As String, dt As String)

    
    Dim newTask As Task
    Set newTask = New Task
    newTask.AddNew
    newTask.TaskName = header
    newTask.Description = content
    newTask.StartDate = DateTime.Date
    newTask.EndDate = dt
    newTask.CFClass = "NONE"
    newTask.Tags = "EMPTY"
    newTask.Status = "PENDING"
    newTask.cardName = lblCardHeader(Index).Caption
    newTask.ProjectId = 1
    newTask.SaveChanges
    
    Call addTile(Index, header, content, dt, "Tile_0")
    
    refreshTaskList
    
    
End Sub


Private Sub lblnewCardHead_Click()
    Dim cardName As String
    cardName = InputBox("Provide a Name for the new card")
    If cardName = "" Then
        Exit Sub
    End If
    Call addCard(pbxCard.Count, cardName)
    ReDim Preserve tileCounts(pbxCard.Count + 1)
End Sub

Private Sub lblNewCardHead_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    SetCursor LoadCursor(0, IDC_HAND)
    lblnewCardHead.ForeColor = RGB(40, 40, 180)
End Sub




Private Sub lblTaskTileContent_DblClick(Index As Integer)
    If TempEditing = True Then
        Exit Sub
    End If
    
    Set txtTempEdit.Container = pbxTaskTile(Index)
    With txtTempEdit
        .Visible = True
        .SetFocus
        .Text = lblTaskTileContent(Index).Caption
        .Left = lblTaskTileContent(Index).Left
        .Top = lblTaskTileContent(Index).Top
        .height = lblTaskTileContent(Index).height
        .width = lblTaskTileContent(Index).width
        .Visible = True
    End With
    TempEditing = True
    While TempEditing
        DoEvents
    Wend
    lblTaskTileContent(Index) = txtTempEdit.Text
    
    
    Dim tTask As Task
    Set tTask = New Task
    tTask.LoadSingleton (Val(pbxTaskTile(Index).tag))
    tTask.Description = txtTempEdit.Text
    tTask.SaveChanges
    
    
    Set txtTempEdit.Container = pbxContent
End Sub

Private Sub lblTaskTileHeader_DblClick(Index As Integer)
    If TempEditing = True Then
        Exit Sub
    End If
    
    Set txtTempEdit.Container = pbxTaskTile(Index)
    With txtTempEdit
        .Visible = True
        .SetFocus
        .Text = lblTaskTileHeader(Index).Caption
        .Left = lblTaskTileHeader(Index).Left
        .Top = lblTaskTileHeader(Index).Top
        .height = lblTaskTileHeader(Index).height
        .width = lblTaskTileHeader(Index).width
    End With
    TempEditing = True
    While TempEditing
        DoEvents
    Wend
    lblTaskTileHeader(Index) = txtTempEdit.Text
    
    
    Dim tTask As Task
    Set tTask = New Task
    tTask.LoadSingleton (Val(pbxTaskTile(Index).tag))
    tTask.TaskName = txtTempEdit.Text
    tTask.SaveChanges
    
    
    Set txtTempEdit.Container = pbxContent
End Sub



Private Sub pbxReset_Click()
    refreshTaskList
End Sub

Private Sub lblResetIcon_Click()
    refreshTaskList
End Sub


Private Sub pbxReset_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    shpResetIcon.BackColor = &H217866
    lblReset.ForeColor = &H217866
    
End Sub

Private Sub pbxReset_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    SetCursor LoadCursor(0, IDC_HAND)
    
End Sub

Private Sub pbxReset_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    shpResetIcon.BackColor = &HE0E0E0
    lblReset.ForeColor = &H111111
    
End Sub

Private Sub lblResetIcon_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    shpResetIcon.BackColor = &H217866
    lblReset.ForeColor = &H217866
    
End Sub

Private Sub lblResetIcon_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    SetCursor LoadCursor(0, IDC_HAND)
    
End Sub

Private Sub lblResetIcon_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    shpResetIcon.BackColor = &HE0E0E0
    lblReset.ForeColor = &H111111
End Sub




Private Sub txtTempEdit_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = vbKeyReturn Then
        TempEditing = False
        txtTempEdit.Visible = False
    End If
    
End Sub


Private Sub pbxCard_DragDrop(Index As Integer, SourceHandler As Control, X As Single, Y As Single)
    Dim Source As Control
    Dim IsSameContainer As Boolean
    IsSameContainer = False
    
    Set Source = SourceHandler.Container
    
    pbxCard(Index).BackColor = preserveBackColor
    pbxContent.BackColor = &HFCFCFC
    lblDeleteNotice.Visible = False
    
    If Source.tag = "Card" Then Exit Sub
    'substract the number of tiles in last card where source was placed
    If Source.Container.tag = "Card" Then
        
        If Index = Source.Container.Index Then
            IsSameContainer = True
            Exit Sub
        End If
        
        'TODO: Implement repositiong of tiles in the card
        ' Now when an element is moved from the tab it is not repositioned back
        ' Steps to reproduce
        ' 1. Add 2 new tasks in a card
        ' 2. drop the first one to card2
        ' 3. drop the tile at card2 back to card1
        ' 4. The tile is now placed at the second place even though there is already another tile
        
        
        If Not IsSameContainer Then
            tileCounts(Source.Container.Index) = tileCounts(Source.Container.Index) - 1
        End If
        
    End If
    
    Call AdjustTileInCard(Index, Source, IsSameContainer)
    '<__________________________
    '
    'Save change to the db
    '___________________________
    
    
    Dim tTask As Task
    Set tTask = New Task
    tTask.LoadSingleton (Val(Source.tag))
    tTask.cardName = lblCardHeader(Index).Caption
    tTask.SaveChanges
    
    refreshTaskList
    
    '__________________________>
    
    
End Sub

Private Sub pbxCard_DragOver(Index As Integer, Source As Control, X As Single, Y As Single, State As Integer)
    
    
    For i = 0 To pbxCard.Count - 1
        If i <> Index Then
            pbxCard(i).BackColor = preserveBackColor
        End If
    Next i
    If Source.tag <> "Card" Then
        pbxCard(Index).BackColor = &H44FF77
        pbxContent.BackColor = &HFCFCFC
        lblDeleteNotice.Visible = False
    Else
        'pbxCard(index).BackColor = &H4444FF
        
    End If
    
End Sub

Private Sub pbxContent_DragDrop(Source As Control, X As Single, Y As Single)
    If Source.tag = "Card" Then
        Source.BackColor = preserveBackColor
        'Source.Left = X
    Else
    
        removeTile Source
        pbxContent.BackColor = &HFCFCFC
        lblDeleteNotice.Visible = False
        pbxCard(Source.Container.Container.Index).BackColor = preserveBackColor
        refreshTaskList
        
    End If
End Sub

Private Sub pbxContent_DragOver(Source As Control, X As Single, Y As Single, State As Integer)
    pbxContent.BackColor = &H8080FF
    lblDeleteNotice.Visible = True
End Sub



Private Sub lblCardAddNew_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    SetCursor LoadCursor(0, IDC_HAND)
    
End Sub
Private Sub lblCardAddNew_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblCardAddNew(Index).ForeColor = RGB(40, 40, 180)
    
End Sub

Private Sub lblCardAddNew_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblCardAddNew(Index).ForeColor = RGB(40, 80, 120)
    
End Sub



Private Sub hscrContent_Change()
    Dim moveLeft As Integer
    moveLeft = -hscrContent.value
    
    If moveLeft < pbxContent.width Then
        pbxContent.Left = moveLeft
    End If
    If pbxCard(pbxCard.Count - 1).Visible Then
        pbxCard(pbxCard.Count - 1).SetFocus
    End If
End Sub


'Functions


'init
Private Sub initTaskView()
    Set cCardManager = Nothing
    Set tCards = Nothing
    Set cCardManager = New Collection
    Set tCards = New TaskCards
    tCards.getCardData

    Dim tempTaskManager As TaskManager
    If pbxCard.Count <= tCards.Count Then
        ReDim Preserve tileCounts(tCards.Count)
        For i = 1 To tCards.Count
            Call addCard(i, tCards.Item(i).cardName)
            
            Set tempTaskManager = New TaskManager
            Call tempTaskManager.loadTasksFromDB(" card_name = '" & tCards.Item(i).cardName & "' ")
            If tempTaskManager.IsLoaded Then
                cCardManager.Add tempTaskManager
                For j = 1 To tempTaskManager.Count
                    
                    Call addTile(i, tempTaskManager.Tasks(j).TaskName, tempTaskManager.Tasks(j).Description, tempTaskManager.Tasks(j).EndDate, tempTaskManager.Tasks(j).TaskId)
                Next j
            End If
        Next i
    End If
End Sub

Private Sub alignContainers()
    
    pbxHead.Top = 0
    pbxHead.Left = 0
    pbxHead.width = Me.width
    pbxHead.height = 1200
    
    pbxContent.Top = pbxHead.Top + pbxHead.height
    pbxContent.Left = 0
    pbxContent.width = Me.width
    pbxContent.height = Me.height - pbxContent.Top - hscrContent.height
    
    pbxReset.Left = pbxHead.width - pbxReset.width - 100
    pbxReset.Top = 10
    
    
End Sub
Private Function CreateCard() As Integer
    
    'load the contents
    Load pbxCard(pbxCard.Count)
    Load lblCardHeader(lblCardHeader.Count)
    Load lblCardAddNew(lblCardAddNew.Count)
    
    Set lblCardHeader(lblCardHeader.Count - 1).Container = pbxCard(pbxCard.Count - 1)
    Set lblCardAddNew(lblCardAddNew.Count - 1).Container = pbxCard(pbxCard.Count - 1)
    
    pbxCard(pbxCard.Count - 1).Visible = True
    lblCardAddNew(lblCardAddNew.Count - 1).Visible = True
    lblCardHeader(lblCardHeader.Count - 1).Visible = True
    
    CreateCard = pbxCard.Count - 1

End Function

Private Function CreateTaskTile() As Integer
    Dim nextIndex As Integer
    nextIndex = pbxTaskTile.Count
    
    'load the contents
    Load pbxTaskTile(nextIndex)
    Load lblTaskTileHeader(nextIndex)
    Load lblTaskTileContent(nextIndex)
    Load lblTaskTileDate(nextIndex)
    Load lblTileHandle(nextIndex)
    
    Set lblTaskTileHeader(lblTaskTileHeader.Count - 1).Container = pbxTaskTile(nextIndex)
    Set lblTaskTileContent(lblTaskTileContent.Count - 1).Container = pbxTaskTile(nextIndex)
    Set lblTaskTileDate(lblTaskTileDate.Count - 1).Container = pbxTaskTile(nextIndex)
    Set lblTileHandle(lblTileHandle.Count - 1).Container = pbxTaskTile(nextIndex)
    
    pbxTaskTile(pbxTaskTile.Count - 1).Visible = True
    lblTaskTileHeader(lblTaskTileHeader.Count - 1).Visible = True
    lblTaskTileContent(lblTaskTileContent.Count - 1).Visible = True
    lblTaskTileDate(lblTaskTileDate.Count - 1).Visible = True
    lblTileHandle(lblTileHandle.Count - 1).Visible = True
    CreateTaskTile = pbxTaskTile.Count - 1

End Function

Private Sub SetTileInfo(Index As Integer, header As String, content As String, dt As String, Optional tag As String = "")
    
    lblTaskTileHeader(Index).Caption = header
    lblTaskTileContent(Index).Caption = content
    lblTaskTileDate(Index).Caption = "  " & dt
    If Len(tag) > 0 Then
        pbxTaskTile(Index).tag = tag
    End If

End Sub

Private Sub AdjustTileInCard(ByVal Index As Integer, Source As Control, SameContainer As Boolean)
    
    Dim baseTop As Integer
    'Make the new tile inside the host Card
    Set Source.Container = pbxCard(Index)
    Source.Left = 360
    
    baseTop = 200 + lblCardHeader(Index).Top + lblCardHeader(Index).height
    If tileCounts(Index) = 0 Then
        Source.Top = baseTop
    Else
        Source.Top = tileCounts(Index) * pbxTaskTile(0).height + tileCounts(Index) * 100 + baseTop
    End If
    If Not SameContainer Then
        'add the number of tiles in current Card
        tileCounts(Index) = tileCounts(Index) + 1
    End If
    
    
    ' TODO: When New tiles are moved or added,
    '   The Card should be auto sized to fit more tiles
    pbxCard(Index).height = tileCounts(Index) * pbxTaskTile(0).height + tileCounts(Index) * 100 + baseTop + lblCardAddNew(Index).height + lblCardHeader(Index).height
    
    lblCardAddNew(Index).Top = pbxCard(Index).height - lblCardAddNew(Index).height
    
End Sub

Private Sub addTile(ByVal Index As Integer, header As String, content As String, dt As String, Optional tag As String = "")
    
    Dim lastIndex As Integer
    lastIndex = CreateTaskTile()
    
    SetTileInfo lastIndex, header, content, dt, tag
    Set pbxTaskTile(lastIndex).Container = pbxCard(Index)
    Call AdjustTileInCard(Index, pbxTaskTile(lastIndex), False)
    
End Sub

Private Sub addCard(ByVal Index As Integer, header As String)
    Dim lastIndex As Integer
    lastIndex = CreateCard()
    Set pbxCard(lastIndex).Container = pbxContent
    If lastIndex > 1 Then
        pbxCard(lastIndex).Top = pbxCard(Index - 1).Top
        pbxCard(lastIndex).Left = pbxCard(Index - 1).Left + pbxCard(Index - 1).width + 560
    Else
        pbxCard(lastIndex).Top = pbxCard(0).Top
        pbxCard(lastIndex).Left = 560
    End If
    
    pbxNewCard.Left = pbxCard(lastIndex).Left + pbxCard(lastIndex).width + 260
    lblCardHeader(Index).Caption = header
    
    Dim sizeOfContent As Double
    Dim maxScroll As Double
    
    sizeOfContent = pbxNewCard.Left + pbxNewCard.width
    
    If pbxContent.width < sizeOfContent Then
        pbxContent.width = sizeOfContent + 560
        
        maxScroll = pbxContent.width - Me.width
        
        If pbxContent.width > Me.width Then
            
            pbxContent.Left = -(maxScroll)
            
            hscrContent.Max = maxScroll
            hscrContent.value = maxScroll
            
            hscrContent.LargeChange = Me.width / 10
            hscrContent.Visible = True
            
            
            hscrContent.Top = Me.height - hscrContent.height
            hscrContent.Left = 0
            hscrContent.width = Me.width
        End If
        
    End If
    
    
    
End Sub


Private Sub removeTile(SourceHandle As Control)
'On Error GoTo WrongPlace:
    Dim Source As Control
    
    Set Source = SourceHandle.Container

    'If TypeName(Source.Container) = "Object" Then
        tileCounts(Source.Container.Index) = tileCounts(Source.Container.Index) - 1
    'End If
    
    'Odd Errors are occured when removing a task
    
    'Do not unload the task for now just hide it
    'TODO: Plan how to remove a task
    
    '_______________________________________
    
    'Unload lblTaskTileContent(Source.Index)
    'Unload lblTaskTileDate(Source.Index)
    'Unload lblTaskTileHeader(Source.Index)
    'Unload pbxTaskTile(Source.Index)
    '_______________________________________
    
    
    pbxCard(Source.Container.Index).height = tileCounts(Source.Container.Index) * pbxTaskTile(0).height + tileCounts(Source.Container.Index) * 100 + 495 + lblCardHeader(Source.Container.Index).height + lblCardAddNew(Source.Container.Index).height
    
    lblCardAddNew(Source.Container.Index).Top = pbxCard(Source.Container.Index).height - lblCardAddNew(Source.Container.Index).height
    
    
    Dim tTask As New Task
    tTask.LoadSingleton (Val(Source.tag))
    tTask.Delete
    tTask.SaveChanges
    pbxTaskTile(Source.Index).Visible = False
    
    Exit Sub
WrongPlace:
    'Some thing went wrong
    MsgBox "Error Occured While deleting Task" & vbLf & Err.Description
End Sub

Private Sub resetControls()
    txtTempEdit.Container = pbxContent
    For i = 1 To pbxTaskTile.Count - 1
        
        Unload lblTaskTileContent(i)
        Unload lblTaskTileDate(i)
        Unload lblTaskTileHeader(i)
        Unload lblTileHandle(i)
        Unload pbxTaskTile(i)
        
    Next i
    
    For i = 1 To pbxCard.Count - 1
        
        Unload lblCardAddNew(i)
        Unload lblCardHeader(i)
        
        Unload pbxCard(i)
        
    Next i
    
    ReDim tileCounts(1)
    
    Screen.MousePointer = vbNormal
End Sub

Private Sub refreshTaskList()

    pbxContent.Visible = False
    Screen.MousePointer = vbHourglass
    resetControls
    initTaskView
    Screen.MousePointer = vbDefault
    pbxContent.Visible = True
End Sub
