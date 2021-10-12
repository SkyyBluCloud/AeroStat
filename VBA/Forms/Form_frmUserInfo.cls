VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmUserInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub cboAgency_AfterUpdate()
    With rankID
        .Value = ""
        .Requery
        .SetFocus
        .Dropdown
    End With
End Sub

Private Sub continue_Click()
Dim errors As Boolean
errors = False
errText.Visible = False

    Dim ctl As Control
    For Each ctl In Controls
        If TypeOf ctl Is TextBox Then
            ctl.BorderColor = RGB(166, 166, 166)
            rankID.BorderColor = RGB(166, 166, 166)
        End If
    Next
    
    For Each ctl In Controls
        If TypeOf ctl Is TextBox Or TypeOf ctl Is ComboBox Then
            If IsNull(ctl) Then
                'If Nz(rankID, 0) = 0 Then rank.BorderColor = RGB(255, 0, 0)
                ctl.BorderColor = RGB(255, 0, 0)
                errText.Caption = "Please enter a value."
                errText.Visible = True
                errors = True
            End If
        End If
    Next
    If errors Then Exit Sub
    
    If Not IsNull(DLookup("ID", "tblUserAuth", "opinitials = '" & opInitials & "'")) Then
        errText.Caption = "These initials have already been assigned. Please choose a different one."
        errText.Visible = True
        opInitials.BorderColor = RGB(255, 0, 0)
        Exit Sub
    End If
    
    opInitials = UCase(opInitials)
    firstName = UCase(Left(firstName, 1)) & LCase(Right(firstName, Len(firstName) - 1))
    lastName = UCase(Left(lastName, 1)) & LCase(Right(lastName, Len(lastName) - 1))
    username = Environ$("username")
    lastLogin = Now()
    lastSystem = Environ$("computername")
    'DoCmd.RunCommand acCmdSaveRecord
    
    ''''''exit access here and constantly check for verification before the main form loads...'''''''''''''''''
    '''For now, we will just continue'''''''''''
    
    isLoggedIn = True
    DoCmd.Close acForm, "frmUserInfo"
    MsgBox "Your information has been saved.", vbInformation, "AeroStat"
    DoCmd.OpenForm "frmMain"
End Sub

Private Sub firstName_AfterUpdate()
firstName = UCase(firstName)
End Sub

Private Sub lastName_AfterUpdate()
lastName = UCase(lastName)
End Sub

Private Sub opInitials_AfterUpdate()
opInitials = UCase(opInitials)
End Sub

Private Sub rankID_AfterUpdate()
lastName.SetFocus
End Sub
