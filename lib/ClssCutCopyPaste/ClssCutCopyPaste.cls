VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ClssCutCopyPaste"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
#If Mac Then
#Else
    Private WithEvents CopyCtl As CommandBarButton
Attribute CopyCtl.VB_VarHelpID = -1
    Private WithEvents CutCtl As CommandBarButton
Attribute CutCtl.VB_VarHelpID = -1
    Private WithEvents PasteCtl As CommandBarButton
Attribute PasteCtl.VB_VarHelpID = -1
    Public WithEvents TxtBox  As msforms.TextBox
Attribute TxtBox.VB_VarHelpID = -1
#End If

Private Sub CopyCtl_Click _
(ByVal Ctrl As Office.CommandBarButton, CancelDefault As Boolean)
    SendKeys ("^{C}")
End Sub

Private Sub CutCtl_Click _
(ByVal Ctrl As Office.CommandBarButton, CancelDefault As Boolean)
    SendKeys ("^{X}")
End Sub

Private Sub PasteCtl_Click _
(ByVal Ctrl As Office.CommandBarButton, CancelDefault As Boolean)
    SendKeys ("^{V}")
End Sub


Private Sub TxtBox_MouseUp _
(ByVal Button As Integer, ByVal Shift As Integer, _
ByVal X As Single, ByVal Y As Single)

    Dim oCmbar As CommandBar
    
    If Button = 2 Then
        Set oCmbar = Application.CommandBars.Add _
        (Position:=msoBarPopup, temporary:=True)
        Set CopyCtl = oCmbar.Controls.Add(msoControlButton)
        Set CutCtl = oCmbar.Controls.Add(msoControlButton)
        Set PasteCtl = oCmbar.Controls.Add(msoControlButton)
        
        With CopyCtl
            .Style = msoButtonIconAndCaption
            .FaceId = 19
            .Caption = "Copy"
        End With
        With CutCtl
            .Style = msoButtonIconAndCaption
            .FaceId = 21
            .Caption = "Cut"
        End With
        With PasteCtl
            .Style = msoButtonIconAndCaption
            .FaceId = 22
            .Caption = "Paste"
        End With
        
        oCmbar.ShowPopup
        oCmbar.Delete
    End If

End Sub

