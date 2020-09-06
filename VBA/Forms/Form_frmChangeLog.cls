VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmChangeLog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Dim bClose As Boolean

Private Sub BOOT_Click()
Dim RS As DAO.Recordset
Set RS = CurrentDb.OpenRecordset("tblUserAuth")
With RS

    If IsNull(bootUser) Then
        If MsgBox("Boot everyone?", vbQuestion + vbYesNo, "BOOT") = vbYes Then
            Do While Not .EOF
                .edit
                !BOOT = True
                .update
                .MoveNext
            Loop
            MsgBox "Done!", vbInformation, "BOOT"
        End If
        
    Else
        Do While Not .EOF
            If !opInitials = bootUser Then
                If Not !isLoggedIn Then
                    MsgBox !rankID & " " & !lastName & " is not logged in.", vbInformation, "BOOT"
                    Exit Do
                ElseIf !BOOT Then
                    MsgBox !rankID & " " & !lastName & " is already being BOOTed.", vbInformation, "BOOT"
                    Exit Do
                End If
                If MsgBox(!rankID & " " & !lastName & vbCrLf & "BOOT?", vbQuestion + vbYesNo, "BOOT") = vbNo Then Exit Do
                
                .edit
                !BOOT = True
                .update
                MsgBox !rankID & " " & !lastName & " has been booted!", vbInformation, "BOOT"
                Exit Do
                
            Else
                .MoveNext
                If .EOF Then MsgBox "User not found", vbInformation, "BOOT"
            End If
        Loop
    End If
    
    .Close
    End With
    Set RS = Nothing
    
End Sub

Private Sub Form_Load()
    version = DMax("version", "tblversioncontrol")
End Sub

Private Sub Form_BeforeUpdate(cancel As Integer)
    If Not bClose Then
        cancel = True
        Undo
        Exit Sub
    End If
End Sub

Private Sub Form_Current()
    bClose = False
End Sub

Private Sub Form_Error(DataErr As Integer, response As Integer)
    If DataErr = 2169 Then response = True
End Sub

Private Sub push_Click()
Dim RS As DAO.Recordset
Dim rsLcl As DAO.Recordset
bClose = True
    If Not IsNull(change) Then
        Set rsLcl = CurrentDb.OpenRecordset("lclVer")
        With rsLcl
            .edit
            !Minor = !Minor + 1
            !Build = !Build + 1
            .update
            .Bookmark = .LastModified
            Set RS = CurrentDb.OpenRecordset("tblSettings")
            RS.edit
            RS!LastVer = !version
            version = RS!LastVer
            RS.update
        End With
        
        'DoCmd.RunCommand acCmdSaveRecord
        MsgBox "Version " & version & " pushed successfully!", vbInformation, "Change Success"
        DoCmd.Close acForm, "frmChangeLog", acSaveYes
    End If
End Sub
