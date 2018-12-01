VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmSnippetView 
   BackColor       =   &H00FFFFFF&
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
   Begin RichTextLib.RichTextBox rtbSnippetView 
      Height          =   4995
      Left            =   315
      TabIndex        =   1
      Top             =   1140
      Width           =   11595
      _ExtentX        =   20452
      _ExtentY        =   8811
      _Version        =   393217
      BorderStyle     =   0
      Enabled         =   -1  'True
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
   Begin VB.PictureBox pbxHead 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H00F7F7F7&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1110
      Left            =   0
      ScaleHeight     =   1110
      ScaleWidth      =   11955
      TabIndex        =   0
      Top             =   0
      Width           =   11955
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
      Height          =   4935
      Left            =   -45
      TabIndex        =   2
      Top             =   1155
      Width           =   435
   End
End
Attribute VB_Name = "frmSnippetView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    alignContainers
End Sub
Private Sub Form_Resize()
    alignContainers
End Sub

Private Sub alignContainers()
    
    'Header PictureBox
    pbxHead.Top = 0
    
    'LineNumberLabel
    lblLineNos.Left = 0
    lblLineNos.height = Me.height - pbxHead.height
    lblLineNos.Top = pbxHead.height + pbxHead.Top
    lblLineNos.width = 360
    
    'RichTextBox
    rtbSnippetView.Left = 360
    
    rtbSnippetView.width = Me.width - lblLineNos.width
    rtbSnippetView.Top = pbxHead.height + pbxHead.Top
    rtbSnippetView.height = Me.height - pbxHead.height
    
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
