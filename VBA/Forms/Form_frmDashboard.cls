VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmDashboard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Dim rst As DAO.Recordset

Private Sub btnAddUser_Click()
DoCmd.OpenForm "frmUserInfo", , , , acFormAdd, acDialog
lstPending.Requery
users.Requery
End Sub

Private Sub btnApprove_Click()
If Nz(lstPending) = "" Then Exit Sub

    CurrentDb.Execute "UPDATE tblUserAuth SET approved = True WHERE ID = " & lstPending
    lstPending.Requery
    users.Requery
    
End Sub

Private Sub btnBOOT_Click()
BOOT Nz(users)
End Sub

Private Sub btnDelete_Click()
If Nz(lstPending) = "" Then Exit Sub

    If MsgBox("This will remove the selected user." & vbCrLf & vbCrLf & _
                "Information related to this user (signatures, log entries, etc) will be kept in the database.", vbOKCancel, "Remove User") = vbOK Then _
        CurrentDb.Execute "DELETE FROM tblUserAuth WHERE ID = " & lstPending
    
End Sub

Private Sub Detail_Click()
    Dim ctl: For Each ctl In Controls
        If TypeOf ctl Is ListBox Then ctl.Value = ""
    Next
End Sub

Private Sub Form_Load()
    Set rst = CurrentDb.OpenRecordset("tblUserAuth")
    With rst
        Do While Not .EOF
            .edit
            If DateDiff("n", !lastLogin, Now) > 2 Then
                !BOOT = False
                !isLoggedIn = False
            End If
            .Update
            .MoveNext
        Loop
    .Close
    End With
    Set rst = Nothing
End Sub

Private Sub Form_Timer()
    users.Requery
End Sub

Public Function BOOT(ByVal users As Integer, Optional RS As Boolean)
Set rst = CurrentDb.OpenRecordset("tblUserAuth")
With rst
.MoveFirst

    If users = 0 Then
        If MsgBox(IIf(RS, "Force reset ", "Boot") & " all users?", vbInformation + vbYesNo, "AeroStat") = vbNo Then Exit Function
        Do While Not .EOF
            .edit
            IIf(RS, !RS, !BOOT) = True
            .Update
            .MoveNext
        Loop
        MsgBox "Booted!", vbInformation, "AeroStat"
    Else
        Do Until !ID = users And Not .EOF: .MoveNext: Loop
        If .EOF Then
            MsgBox "User not found.", vbInformation, "AeroStat"
            Exit Function
        End If
        
        If MsgBox(IIf(RS, "Force reset ", "Boot ") & !rankID & " " & !lastName & "?", vbInformation + vbYesNo, "AeroStat") = vbNo Then Exit Function
        
        .edit
        IIf(RS, !RS, !BOOT) = True
        .Update
        MsgBox "Done!", vbInformation, "AeroStat"
    End If
    .Close
    
End With
Set rst = Nothing
End Function

Private Sub lstPending_Click()
users = ""
End Sub

Private Sub restart_Click()
BOOT Nz(users), True
End Sub

Private Sub users_Click()
'lstPending.Click = ""
End Sub
