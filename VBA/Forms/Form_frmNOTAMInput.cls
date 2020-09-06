VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmNOTAMInput"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Dim bClose As Boolean
Dim shiftID As Integer

Private Sub bParse_Click()
If Nz(parse) = "" Then Exit Sub
'n = NOTAMUtil.parseNOTAM(parse, DLookup("opinitials", "tbluserauth", "username = '" & Environ$("username") & "'"))
N = NOTAMUtil.parseNOTAM(parse)
    If N <> 0 Then
        Undo
        bClose = True
        Me.filter = "ID = " & N
        Me.FilterOn = True
        issuedBy.BorderColor = RGB(255, 0, 0)
    Else
        MsgBox "Could not parse NOTAM. Invalid format.", vbInformation, "NOTAM Parser"
    End If
End Sub

Private Sub Form_Current()
    If NewRecord Then
        parse.Visible = NewRecord
        bParse.Visible = NewRecord
        parse.SetFocus
    End If
    bClose = False
End Sub

Private Sub Form_BeforeUpdate(cancel As Integer)
    If Not bClose Then
        cancel = True
        Undo
        Exit Sub
    End If
End Sub

Private Sub Form_Load()
If IsNull(OpenArgs) Then
    If shiftID = 0 Then
        If Not DLookup("closed", "tblshiftmanager", "shiftid = " & DLookup("lastshift", "tbluserauth", "username = '" & Environ$("username") & "'")) Then
            shiftID = DLookup("lastshift", "tbluserauth", "username = '" & Environ$("username") & "'")
        End If
    End If
Else
    shiftID = OpenArgs
End If
End Sub

Private Sub save_Click()
Dim log As DAO.Recordset
Dim rNOTAM As DAO.Recordset
bClose = True
Dim fail As Boolean
For Each ctl In Controls
    If ctl.Tag = "req" Then
        If Nz(ctl.Value) = "" Then
            ctl.BorderColor = RGB(255, 0, 0)
            fail = True
        Else
            ctl.BorderColor = RGB(166, 166, 166)
        End If
    End If
    If TypeOf ctl Is TextBox Then ctl.Value = UCase(ctl.Value)
Next
If fail Then Exit Sub

If nType = "C" Then NOTAMUtil.cancelNOTAM Right(verbiage, 8), endTime
If Dirty Then Dirty = False

    If logentry And Nz(shiftID) <> "" Then
        Set log = CurrentDb.OpenRecordset("SELECT * FROM tbl3616 WHERE shiftID = " & shiftID)
        With log
            .AddNew
            !shiftID = shiftID
            !entryTime = LToZ(Now)
            Select Case nType
                Case "N", "R"
                    !entry = "NEW NOTAM: " & NOTAM & " - " & verbiage & ". VALID " & Format(startTime, "yymmddhhnn") & " - " & Format(endTime, "yymmddhhnn")
                Case "C"
                    !entry = "NOTAM " & Right(verbiage, 8) & " CANCELLED"
            End Select
            !Initials = DLookup("opinitials", "tbluserauth", "username = '" & Environ$("username") & "'")
            .update
            .Close
        End With
        Set log = Nothing
    End If
            
DoCmd.Close acForm, "frmNOTAMInput", acSaveYes
End Sub
