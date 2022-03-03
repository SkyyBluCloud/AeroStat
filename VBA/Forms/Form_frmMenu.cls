VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub btnContinue_Click()
DoCmd.Close acForm, "frmMenu", acSaveNo
DoCmd.OpenForm "frmMain"
End Sub

Private Sub btnJoinShift_Click()
Dim user As String
user = Util.getUser
    If Not Util.isShiftClosed(user) Then
        If MsgBox("Your previous wasn't closed. Continue last shift?", vbQuestion + vbYesNo, "New Shift") = vbNo Then
            DoCmd.OpenForm "frmShiftSelect", , , , , acDialog
            If Util.isShiftClosed(user) Then
                Exit Sub
            End If
        End If
        DoCmd.Close acForm, "frmMenu", acSaveNo
        DoCmd.OpenForm "frmMain"
    End If
End Sub

Private Sub btnNewShift_Click()

    If Not Util.isShiftClosed(user) Then
        If MsgBox("Your previous shift must be closed before joining a new one. Would you like to review it?", vbQuestion + vbYesNo, "New Shift") = vbNo Then Exit Sub
        DoCmd.OpenForm "frmShiftDash", , , , , acDialog, DLookup("lastshift", "tbluserauth", "username = '" & Util.getUser & "'")
        
    Else
        DoCmd.OpenForm "frmShiftDash", , , , acFormAdd, acDialog, Me.Name
    End If
End Sub

Private Sub Form_Load()
btnContinue.Visible = Not Util.isShiftClosed(user)
End Sub
