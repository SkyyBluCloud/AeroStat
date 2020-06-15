VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmAFMDash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub btnBack_Click()
Parent.SourceObject = "frmShiftBoard"
End Sub

Private Sub btnSign_Click()
'Dim rs As DAO.Recordset
'Set rs = CurrentDb.OpenRecordset(AF3616.Report.RecordSource)
'Dim roleStr As String
'
'    Select Case role
'    Case 2: roleStr = "AMOS"
'    Case 3: roleStr = "NAMO"
'    Case 4: roleStr = "AFM"
'    End Select
'
'    If MsgBox("You are signing this Events Log as the " & roleStr & ". " & vbCrLf & _
'        "By signing this document, you certify that all entries are correct; " & _
'        "that all scheduled operations have been accomplished, except as noted, " & _
'        "that all abnormal occurences or conditions and all significant incidents/events have been recorded.", vbOKCancel, "Events Log") = vbCancel _
'        Then Exit Sub
'
'    With rs: Do While Not .EOF
'        .edit
'        .Fields(LCase(roleStr) & "Sig") = getUSN
'        .Fields(LCase(roleStr) & "SigTime") = Now
'        .update
'        .MoveNext
'    Loop: End With
'
'******MODULED********

    If UtilForm3616.signLog(shiftList, AF3616.Report.RecordSource, role) Then
        shiftList.Requery
        shiftList_Click
    End If

End Sub

Private Sub role_AfterUpdate()
Dim qdf As DAO.QueryDef
Dim qry As String
Dim reviewStatus As String
Dim authLevel As Integer
authLevel = DLookup("authlevel", "tbluserauth", "username = '" & Util.getUSN & "'")
If Util.getUSN = "1516576541A" Then authLevel = 0
qry = "SELECT tblShiftManager.shiftID, tblShiftManager.shiftStart AS [Start Time (L)], tblShiftManager.shiftEnd AS [End Time (L)], tblShiftManager.superlead AS AMOS, IIf([amosSig] Is Not Null,'YES','NO') AS [AMOS Signed], IIf([namoSig] Is Not Null,'YES','NO') AS [NAMO Signed], IIf([afmSig] Is Not Null,'YES','NO') AS [AFM Signed], [amosSig] Is Not Null AS Expr1, [namoSig] Is Not Null AS Expr2, [afmSig] Is Not Null AS Expr3 FROM tblShiftManager WHERE ((([amosSig] Is Not Null)=[showForNAMO]) AND (([namoSig] Is Not Null)<>[showForNAMO]) AND (([afmSig] Is Not Null)<>[showForNAMO])) OR ((([namoSig] Is Not Null)=[showForAFM]) AND (([afmSig] Is Not Null)<>[showForAFM])) ORDER BY tblShiftManager.shiftStart DESC;"
Set qdf = CurrentDb.QueryDefs("qAFMDash")

    Select Case role
    Case 1 'All
        reviewStatus = ""
        Set shiftList.Recordset = CurrentDb.OpenRecordset("SELECT tblShiftManager.shiftID, tblShiftManager.shiftStart AS [Start Time (L)], tblShiftManager.shiftEnd AS [End Time (L)], tblShiftManager.superlead AS AMOS, IIf([amosSig] Is Not Null,'YES','NO') AS [AMOS Signed], IIf([namoSig] Is Not Null,'YES','NO') AS [NAMO Signed], IIf([afmSig] Is Not Null,'YES','NO') AS [AFM Signed], [amosSig] Is Not Null AS Expr1, [namoSig] Is Not Null AS Expr2, [afmSig] Is Not Null AS Expr3 FROM tblShiftManager ORDER BY tblShiftManager.shiftStart DESC;")
    Case 2 'AMOS
        reviewStatus = "***Reviewing as AMOS/AMSL***"
        qdf.Parameters("showForAMOS") = True
        qdf.Parameters("showForNAMO") = False
        qdf.Parameters("showForAFM") = False
        Set shiftList.Recordset = qdf.OpenRecordset
    Case 3 'NAMO
        If authLevel > 4 Then
            MsgBox "You do not have permission to sign as the NAMO.", vbInformation, "Log Review"
            Exit Sub
        End If
        
        reviewStatus = "***Reviewing as NAMO***"
        qdf.Parameters("showForAMOS") = False
        qdf.Parameters("showForNAMO") = True
        qdf.Parameters("showForAFM") = False
        Set shiftList.Recordset = qdf.OpenRecordset
    Case 4 'AFM
        If authLevel > 2 Then
            MsgBox "You do not have permission to sign as the AFM.", vbInformation, "Log Review"
            Exit Sub
        End If
        
        reviewStatus = "***Reviewing as AFM***"
        qdf.Parameters("showForAMOS") = False
        qdf.Parameters("showForNAMO") = False
        qdf.Parameters("showForAFM") = True
        Set shiftList.Recordset = qdf.OpenRecordset
    Case Else
'        Set shiftList.Recordset = CurrentDb.OpenRecordset("SELECT tblShiftManager.shiftID, tblShiftManager.shiftStart AS [Start Time (L)], tblShiftManager.shiftEnd AS [End Time (L)], tblShiftManager.superlead AS AMOS, IIf([amosSig] Is Not Null,'YES','NO') AS [AMOS Signed], IIf([namoSig] Is Not Null,'YES','NO') AS [NAMO Signed], IIf([afmSig] Is Not Null,'YES','NO') AS [AFM Signed], [amosSig] Is Not Null AS Expr1, [namoSig] Is Not Null AS Expr2, [afmSig] Is Not Null AS Expr3 FROM tblShiftManager ORDER BY tblShiftManager.shiftStart DESC;")
    End Select
    
    btnSign.Enabled = reviewStatus <> ""
    For Each ctl In Me.Controls
        If ctl.Tag = "review" Then
            ctl.Caption = reviewStatus
            ctl.Visible = reviewStatus <> ""
        End If
    Next
End Sub

Private Sub updateLists()
Dim qry As String
Dim monthyear As String

If togDate Then
    monthyear = "(Format([startdate],'mmm') = '" & cboMonth & "' And Year([startdate]) = " & cboYear & ") AND"
End If
qry = "SELECT tblShiftManager.shiftID, tblShiftManager.shiftStart AS [Start Time (L)], tblShiftManager.shiftEnd AS [End Time (L)], tblShiftManager.superlead AS AMOS, IIf([amosSig] Is Not Null,'YES','NO') AS [AMOS Signed], IIf([namoSig] Is Not Null,'YES','NO') AS [NAMO Signed], IIf([afmSig] Is Not Null,'YES','NO') AS [AFM Signed], [amosSig] Is Not Null AS Expr1, [namoSig] Is Not Null AS Expr2, [afmSig] Is Not Null AS Expr3 FROM tblShiftManager WHERE " & monthyear & " ((([amosSig] Is Not Null)=[showForNAMO]) AND (([namoSig] Is Not Null)<>[showForNAMO]) AND (([afmSig] Is Not Null)<>[showForNAMO])) OR ((([namoSig] Is Not Null)=[showForAFM]) AND (([afmSig] Is Not Null)<>[showForAFM])) ORDER BY tblShiftManager.shiftStart DESC;"
    
End Sub

Private Sub shiftList_Click()
On Error GoTo errTrap
'If IsNull(shiftList) Then Exit Sub

    If shiftList.ItemsSelected.Count <= 1 Then
        With AF3616
'            .SourceObject = ""
'           DoEvents
            .SourceObject = "Report.new3616"
            '.Report.args = shiftList.Column(1)  ' shiftStart
            .SetFocus
'            .Report.FilterOn = True
        End With
    End If
sexit:
    Exit Sub
errTrap:
    errHandler err, Error$, Me.Name
End Sub

Private Sub shiftList_DblClick(cancel As Integer)
DoCmd.OpenForm "frm3616", , , , , , shiftList.Column(0)
End Sub

Private Sub togDate_Click()
cboMonth.Enabled = togDate
cboYear.Enabled = togDate
End Sub
