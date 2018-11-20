Attribute VB_Name = "Util"
Public Sub UtilCenterForm(frmSub As Form, frmParent As Form)
    frmSub.Move (frmParent.width - frmSub.width) / 2, (frmParent.height - frmSub.height - 800) / 2
End Sub

