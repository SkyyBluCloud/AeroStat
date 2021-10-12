VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmShiftSelect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub join_Click()
On Error GoTo errtrap
Dim ls As Integer
ls = DLookup("lastShift", "tblUserAuth", "username = '" & Environ$("username") & "'")
If IsNull(listShift.Column(0)) Then Exit Sub
If ls = listShift.Column(0) Or Not DLookup("closed", "tblShiftManager", "shiftID = " & ls) Then
    If MsgBox("You have already joined a shift that was not closed. Continue?", vbQuestion + vbYesNo, "AeroStat") = vbNo Then Exit Sub
End If
        
    'appendUser Environ$("username"), "lastShift", listShift.Column(0)
    CurrentDb.Execute "UPDATE tblUserAuth, tblShiftManager " & _
                        "SET tblUserAuth.lastShift = " & listShift.Column(0) & ", " & _
                            "tblUserAuth.onshift = True, " & _
                            "tblShiftManager.amoc = iif(tblShiftManager.amoc Is Not Null,tblShiftManager.amoc & ';','') & '" & getLogName(getOpInitials(Environ("username"))) & _
                            "' WHERE tblUserAuth.username = '" & Environ("username") & "' AND tblShiftManager.shiftID = " & listShift.Column(0)
    
    MsgBox "Shift: " & listShift.Column(3) & vbCrLf & _
    "Start: " & listShift.Column(1) & vbCrLf & _
    "End: " & listShift.Column(2) & vbCrLf & vbCrLf & _
    "Join Successful."
    
    DoCmd.OpenForm "frmShiftBoard", , , "shiftID = " & listShift.Column(0), acFormEdit, , listShift.Column(0)
    DoCmd.Close acForm, "frmShiftSelect"
    
sexit:
    Exit Sub
errtrap:
    MsgBox Error$
    ErrHandler err, Error$, Me.Name
    Resume Next
    
End Sub

Private Sub new_Click()
On Error Resume Next
DoCmd.Close acForm, "frmShiftBoard"
DoCmd.Close acForm, "frmShiftSelect"
DoCmd.OpenForm "frmShiftBoard", , , , acFormAdd
End Sub

Private Sub review_Click()
DoCmd.OpenForm "frmAfmDash"
DoCmd.Close acForm, "frmShiftSelect"
End Sub
