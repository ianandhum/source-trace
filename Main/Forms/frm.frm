VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frm 
   BackColor       =   &H00EEEEEE&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   5565
   ClientLeft      =   5175
   ClientTop       =   2820
   ClientWidth     =   11145
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5565
   ScaleWidth      =   11145
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox Text1 
      BorderStyle     =   0  'None
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   1
      EndProperty
      Height          =   375
      Left            =   4680
      TabIndex        =   4
      Text            =   "Text1"
      Top             =   960
      Width           =   1575
   End
   Begin ComctlLib.TreeView TreeView1 
      Height          =   5535
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   9763
      _Version        =   327682
      LabelEdit       =   1
      Style           =   7
      Appearance      =   0
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
   Begin ComctlLib.ProgressBar ProgressBar1 
      Height          =   495
      Left            =   3720
      TabIndex        =   3
      Top             =   2280
      Width           =   5175
      _ExtentX        =   9128
      _ExtentY        =   873
      _Version        =   327682
      Appearance      =   1
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Cancel"
      Height          =   495
      Left            =   7560
      TabIndex        =   1
      Top             =   4680
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Save"
      Height          =   495
      Left            =   9240
      TabIndex        =   0
      Top             =   4680
      Width           =   1335
   End
   Begin ComctlLib.ImageList ImageList1 
      Left            =   3360
      Top             =   1320
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   16777215
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   3
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frm.frx":0000
            Key             =   "Leaf"
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frm.frx":005E
            Key             =   "Open"
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frm.frx":00BC
            Key             =   "Closed"
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Projects As ProjectManager

Private Sub Form_Load()
    InitializeConnection
End Sub

