VERSION 5.00
Begin VB.UserControl TileIcon 
   Alignable       =   -1  'True
   AutoRedraw      =   -1  'True
   BackStyle       =   0  'Transparent
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3015
   ClipBehavior    =   0  'None
   DrawWidth       =   2
   EditAtDesignTime=   -1  'True
   PaletteMode     =   0  'Halftone
   ScaleHeight     =   3600
   ScaleWidth      =   3015
   Begin VB.Image pbxIcon 
      Height          =   2655
      Left            =   120
      Stretch         =   -1  'True
      Top             =   120
      Width           =   2775
   End
   Begin VB.Label lblText 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   150
      TabIndex        =   0
      Top             =   2880
      Width           =   2715
   End
End
Attribute VB_Name = "TileIcon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public Event Click()

Private mPicture As IPictureDisp
Private mCaptionText As String


Property Get CaptionText() As String
 CaptionText = mCaptionText
End Property


Property Let CaptionText(newCaptionText As String)
 mCaptionText = newCaptionText
 lblText.Caption = mCaptionText
 PropertyChanged "CaptionText"
End Property




Private Sub UserControl_Click()
    RaiseEvent Click
End Sub

Private Sub UserControl_Resize()
Adjust_Position
End Sub

Private Sub Adjust_Position()
With pbxIcon
    .Top = 10
    .Left = 10
    .width = UserControl.width - 20
    .height = UserControl.height - 500
    
    
End With

With lblText
    .Left = 10
    .width = UserControl.width - 20
    .Top = pbxIcon.height + 50
    .height = UserControl.height - pbxIcon.height - 20
End With

End Sub





Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
 Set Picture = PropBag.ReadProperty("Picture", Nothing)
 CaptionText = PropBag.ReadProperty("CaptionText", "")
End Sub


Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
 '
 '
 PropBag.WriteProperty "Picture", Picture, Nothing
 PropBag.WriteProperty "CaptionText", CaptionText, ""
 '
 '
End Sub

Property Set Picture(newPic As IPictureDisp)
 Set mPicture = newPic
 Set pbxIcon.Picture = mPicture
 PropertyChanged "Picture"
End Property

Property Get Picture() As IPictureDisp
 Set Picture = mPicture
End Property

Property Let Picture(newImage As IPictureDisp)
 Set pbxIcon.Picture = newImage
End Property


