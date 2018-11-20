Attribute VB_Name = "MiscCompact"
'The following is code is from different sources and used  as  workaround for VB specifc issues
'Appropriate Credits will be provided for each part


'Code From BruceG@vbforums.com
'Fix to Treeview Background Behaviour

Option Explicit

Private Const GWL_STYLE                  As Long = (-16)
Private Const TVS_HASLINES               As Long = 2
Private Const TV_FIRST                   As Long = &H1100
Private Const TVM_SETBKCOLOR             As Long = (TV_FIRST + 29)

Private Declare Function SendMessage _
   Lib "user32" Alias "SendMessageA" _
   (ByVal hwnd As Long, _
    ByVal wMsg As Long, _
    ByVal wParam As Long, _
    lParam As Any) As Long

Private Declare Function GetWindowLong _
   Lib "user32" Alias "GetWindowLongA" _
   (ByVal hwnd As Long, _
    ByVal nIndex As Long) As Long

Private Declare Function SetWindowLong _
   Lib "user32" Alias "SetWindowLongA" _
   (ByVal hwnd As Long, _
    ByVal nIndex As Long, _
    ByVal dwNewLong As Long) As Long

'-----------------------------------------------------------------------------
Public Sub SetTreeViewBackColor(pobjTV As TreeView, plngBackColor As Long)
'-----------------------------------------------------------------------------

   Dim lngTVHwnd   As Long
   Dim lngStyle    As Long
   Dim objTVNode   As Node
   
   lngTVHwnd = pobjTV.hwnd
   
   ' Change the background
   Call SendMessage(lngTVHwnd, TVM_SETBKCOLOR, 0, ByVal plngBackColor)
   

   ' Reset the treeview style so the tree lines appear properly ...
   lngStyle = GetWindowLong(lngTVHwnd, GWL_STYLE)
   
   ' If the treeview has lines, temporarily remove them so the back
   ' repaints to the selected colour, then restore ...
   If lngStyle And TVS_HASLINES Then
      Call SetWindowLong(lngTVHwnd, GWL_STYLE, lngStyle Xor TVS_HASLINES)
      Call SetWindowLong(lngTVHwnd, GWL_STYLE, lngStyle)
   End If
   
End Sub


