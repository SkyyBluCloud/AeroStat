VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmInspection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Dim bClose As Boolean

Private Sub add483_Click()
    DoCmd.OpenForm "frm483CompCheck", , , , acFormAdd, acDialog, shiftID
    Remarks.Requery
End Sub

Private Sub done_Click()
On Error GoTo errtrap
bClose = True
Dim rsLog As DAO.Recordset
Dim entry As String
Dim t As Date
    
'    If IsNull(outTime) Then
'        outTime.BorderColor = RGB(255, 0, 0)
'        If IsNull(inTime) Then
'            If Not NewRecord Then
'                inTime.BorderColor = RGB(255, 0, 0)
'                Exit Sub
'            End If
'        End If
'        Exit Sub
'    End If
    
    Dim fail As Boolean
    For Each ctl In Controls
        If TypeOf ctl Is TextBox Then
            Select Case ctl.Name
                Case "callsign", "initials", "outtime"
                    If IsNull(ctl.Value) Then
                        ctl.BorderColor = RGB(255, 0, 0)
                        fail = True
                    Else
                        ctl.BorderColor = RGB(166, 166, 166)
                    End If
                
                Case "intime"
                    If IsNull(ctl.Value) And Not NewRecord Then
                        ctl.BorderColor = RGB(255, 0, 0)
                        fail = True
                    Else
                        ctl.BorderColor = RGB(166, 166, 166)
                    End If
            End Select
        End If
    Next
    If fail Then Exit Sub
    
    If NewRecord Then
        entry = Callsign & "/" & Initials & " ON AFLD FOR AFLD CK"
        t = checkDate & " " & outTime
    Else
        entry = Callsign & "/" & Initials & " OFF AFLD. " & IIf(IsNull(Remarks), "NO CHANGES TO REPORT", Remarks)
        t = checkDate & " " & inTime
    End If
    DoCmd.RunCommand acCmdSaveRecord

    If MsgBox("Add log entry?", vbQuestion + vbYesNo, "Inspection") = vbYes Then
        Set rsLog = CurrentDb.OpenRecordset("SELECT * FROM tbl3616 WHERE shiftid = " & shiftID)
log:    With rsLog
            If MsgBox("Add log entry?", vbQuestion + vbYesNo, "483 Check") = vbYes Then
                .AddNew
                !shiftID = shiftID
                !entryTime = t
                !entry = entry
                !Initials = DLookup("opinitials", "tbluserauth", "username = '" & Environ$("username") & "'")
                .Update
            End If
        End With
        If Not IsNull(inTime) Then
            entry = Callsign & "/" & Initials & " OFF AFLD. " & IIf(IsNull(Remarks), "NO CHANGES TO REPORT", Remarks)
            t = checkDate & " " & inTime
            GoTo log
        End If
    End If
sExit:
    DoCmd.Close acForm, "frm483CompCheck", acSaveYes
    Exit Sub
errtrap:
    If err = 3101 Then
        MsgBox "This 483 is not in the database.", vbInformation, "AF483"
    End If
End Sub

Private Sub Form_Load()
    If IsNull(outTime) Then
        outTime = LToZ(Now)
    ElseIf IsNull(inTime) Then
        inTime = LToZ(Now)
    End If
    shiftID = OpenArgs
End Sub

Private Sub Inspection_Click()
    checklist.Visible = Inspection
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
    If DataErr = 2169 Or DataErr = 3101 Then response = True
End Sub

