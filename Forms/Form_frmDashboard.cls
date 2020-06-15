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

Private Sub btnBOOT_Click()
BOOT Nz(users)
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

Public Function BOOT(ByVal users As Integer, Optional rs As Boolean)
Set rst = CurrentDb.OpenRecordset("tblUserAuth")
With rst
.MoveFirst

    If users = 0 Then
        If MsgBox(IIf(rs, "Force reset ", "Boot") & " all users?", vbInformation + vbYesNo, "AeroStat") = vbNo Then Exit Function
        Do While Not .EOF
            .edit
            IIf(rs, !rs, !BOOT) = True
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
        If MsgBox(IIf(rs, "Force reset ", "Boot ") & !rank & " " & !lastName & "?", vbInformation + vbYesNo, "AeroStat") = vbNo Then Exit Function
        .edit
        IIf(rs, !rs, !BOOT) = True
        .Update
        MsgBox "Done!", vbInformation, "AeroStat"
    End If
    .Close
    
End With
Set rst = Nothing
End Function

Private Sub restart_Click()
BOOT Nz(users), True
End Sub
