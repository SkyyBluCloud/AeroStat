VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmFeedback"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub cancel_Click()
    DoCmd.Close acForm, "frmFeedback"
End Sub

Private Sub okay_Click()
    If IsNull(Feedback) Then
        MsgBox "No blank feedback please...", vbCritical, "Write Something"
    Else
        DoCmd.RunCommand acCmdSaveRecord
        DoCmd.Close acForm, "frmFeedback"
        MsgBox "Thanks for the feedback! :D", vbInformation, "Thank You"
    End If
End Sub
