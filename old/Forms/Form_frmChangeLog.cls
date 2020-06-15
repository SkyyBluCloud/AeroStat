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
Dim rs As DAO.Recordset
Set rs = CurrentDb.OpenRecordset("tblUserAuth")
With rs

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
                    MsgBox !rank & " " & !lastName & " is not logged in.", vbInformation, "BOOT"
                    Exit Do
                ElseIf !BOOT Then
                    MsgBox !rank & " " & !lastName & " is already being BOOTed.", vbInformation, "BOOT"
                    Exit Do
                End If
                If MsgBox(!rank & " " & !lastName & vbCrLf & "BOOT?", vbQuestion + vbYesNo, "BOOT") = vbNo Then Exit Do
                
                .edit
                !BOOT = True
                .update
                MsgBox !rank & " " & !lastName & " has been booted!", vbInformation, "BOOT"
                Exit Do
                
            Else
                .MoveNext
                If .EOF Then MsgBox "User not found", vbInformation, "BOOT"
            End If
        Loop
    End If
    
    .Close
    End With
    Set rs = Nothing
    
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

Private Sub Form_Error(DataErr As Integer, Response As Integer)
    If DataErr = 2169 Then Response = True
End Sub

Private Sub push_Click()
Dim rs As DAO.Recordset
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
            Set rs = CurrentDb.OpenRecordset("settings")
            rs.edit
            rs!LastVer = !version
            version = rs!LastVer
            rs.update
        End With
        
        'DoCmd.RunCommand acCmdSaveRecord
        MsgBox "Version " & version & " pushed successfully!", vbInformation, "Change Success"
        DoCmd.Close acForm, "frmChangeLog", acSaveYes
    End If
End Sub
