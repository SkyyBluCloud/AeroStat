VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmReports"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private currentReportDate As Variant

Public Property Let reportDate(d As Variant)
    currentReportDate = DateValue(d)
    frmReportsView.Form.initializeSubForms
    frmReportsView.SetFocus
End Property

Public Property Get reportDate() As Variant
    If IsEmpty(currentReportDate) Then currentReportDate = Date
    reportDate = currentReportDate
End Property

Private Sub btnUnsign_Click()
If IsNull(shiftList.Column(0)) Then Exit Sub
    
    If role = 3 Then
        If MsgBox("This will also remove the NAMO/AMOM signature. Continue?", vbQuestion + vbYesNo, "Unsign") = vbNo Then Exit Sub
    End If
        
    If Not UtilCertifier.deCertifyDay(role, getUSN, shiftList.Column(0)) Then
        MsgBox "Could not remove signatures.", vbInformation, "Reports"
    Else
        frmReportsView.Form.initializeSubForms
        shiftList.Requery
        frmReportsView.SetFocus
    End If
End Sub

Private Sub Form_Load()
    Dim qdf As DAO.QueryDef
    Set qdf = CurrentDb.QueryDefs("qDailyReportIndex")
    qdf.Parameters("varSigned") = True
    qdf.Parameters("varConcur") = True
    qdf.Parameters("varCert") = True
    
    Set shiftList.Recordset = qdf.OpenRecordset
    frmReportsView.SetFocus
End Sub

Private Sub shiftList_Click()
If Nz(shiftList.Column(0)) = "" Then Exit Sub
With shiftList

    reportDate = .Column(0)
    If role <> 1 Then
            
        Dim sIDX: For sIDX = -2 To 2 'Check previous and next two rows for matching dates, then select them.
            Do While .ItemsSelected(0) + sIDX <= 0:  sIDX = sIDX + 1: Loop
            If .ItemsSelected(0) + sIDX > .ListCount - 1 Then Exit For
            
            If DateValue(.Column(0, .ItemsSelected(0) + sIDX)) = DateValue(reportDate) Then
                .Selected(.ItemsSelected(0) + sIDX) = True
            End If
        Next sIDX
        'Next idx
    End If
    DoEvents
End With
End Sub

'Private Sub btnComment_Click()
'    Dim res As String
'    res = InputBox("Enter your comment for this shift:" & vbCrLf & vbCrLf & Format(shiftList.Column(1), "dd-mmm-yy hh:nnL") & " - " & Format(shiftList.Column(2), "dd-mmm-yy hh:nnL") & ": " & vbCrLf & shiftList.Column(3), _
'                "Reviewer Comments", Nz(DLookup("reviewerComments", "tblShiftManager", "shiftID=" & shiftList.Column(1))))
'    If Not IsNull(res) Then CurrentDb.Execute "UPDATE tblShiftManager SET reviewerComments = '" & res & "' WHERE shiftID = " & shiftList.Column(1), dbFailOnError
'    frmReportsView.Form.AF3616.Requery
'End Sub

Private Sub btnSign_Click()
Dim certifier As Double
If IsNull(shiftList.Column(0)) Then Exit Sub
    
    'TODO: Don't hard-code
    If MsgBox("By signing this document, you certify that all entries are correct; " & _
        "that all scheduled operations have been accomplished, except as noted; " & _
        "that all abnormal occurences or conditions and all significant incidents/events have been recorded.", vbOKCancel + vbInformation, "Events Log") = vbCancel _
    Then Exit Sub
    
    Select Case role
    Case 1
        certifier = UtilCertifier.newCert(getUSN)
        
        If certifier <> 0 Then
            CurrentDb.Execute "UPDATE tblShiftManager SET certifierID = " & certifier & " WHERE shiftID = " & shiftList.Column(1)
            
            log "Success!", Me.Name & ".btnSign_Click"
            Call role_AfterUpdate
        Else
            log "Could not sign report. Cert " & certifier, Me.Name & ".btnSign_Click", "WARN"
            GoTo sexit
        End If
        
    Case 2, 3 'NAMO or AFM
        If Not UtilCertifier.certifyShiftDay(role, getUSN, shiftList.Column(0)) Then
            log "Could not sign report. Cert = " & certifier, Me.Name & "btnSign_Click", "WARN"
            GoTo sexit
        End If
        
    End Select
    frmReportsView.Form.initializeSubForms
    shiftList.Requery
    frmReportsView.SetFocus
        
sexit:
    Exit Sub
    
errtrap:
    ErrHandler err, Error$, "frmReports.btnSign_Click"
    
End Sub

Private Sub role_AfterUpdate()
Dim qdf As DAO.QueryDef
Dim reviewStatus As String
Dim authLevel As Integer
authLevel = DLookup("authlevel", "tbluserauth", "username = '" & Util.getUSN & "'")
If CurrentProject.Name Like "*DEV*" Then authLevel = 0

Set qdf = CurrentDb.QueryDefs("qDailyReportIndex")

    Select Case role
    Case 0 'All
        reviewStatus = ""
        qdf.Parameters("varSigned") = True
        qdf.Parameters("varConcur") = True
        qdf.Parameters("varCert") = True
        Set shiftList.Recordset = qdf.OpenRecordset
        
    Case 1 'AMOS
        reviewStatus = "***Reviewing as AMOS/AMSL***"
        qdf.Parameters("varSigned") = False
        qdf.Parameters("varConcur") = False
        qdf.Parameters("varCert") = False
        Set shiftList.Recordset = qdf.OpenRecordset
        
    Case 2 'NAMO
'        If authLevel > 4 Then
'            MsgBox "You do not have permission to sign as the NAMO.", vbInformation, "Log Review"
'            Exit Sub
'        End If
        
        reviewStatus = "***Reviewing as NAMO/AMOM***"
        qdf.Parameters("varSigned") = True
        qdf.Parameters("varConcur") = False
        qdf.Parameters("varCert") = False
        Set shiftList.Recordset = qdf.OpenRecordset
        
    Case 3 'AFM
'        If authLevel > 2 Then
'            MsgBox "You do not have permission to sign as the AFM.", vbInformation, "Log Review"
'            Exit Sub
'        End If
        
        reviewStatus = "***Reviewing as AFM***"
        qdf.Parameters("varSigned") = True
        qdf.Parameters("varConcur") = True
        qdf.Parameters("varCert") = False
        Set shiftList.Recordset = qdf.OpenRecordset
    Case Else
'        Set shiftList.Recordset = CurrentDb.OpenRecordset("SELECT tblShiftManager.shiftID, tblShiftManager.shiftStart AS [Start Time (L)], tblShiftManager.shiftEnd AS [End Time (L)], tblShiftManager.superlead AS AMOS, IIf([amosSig] Is Not Null,'YES','NO') AS [AMOS Signed], IIf([namoSig] Is Not Null,'YES','NO') AS [NAMO Signed], IIf([afmSig] Is Not Null,'YES','NO') AS [AFM Signed], [amosSig] Is Not Null AS Expr1, [namoSig] Is Not Null AS Expr2, [afmSig] Is Not Null AS Expr3 FROM tblShiftManager ORDER BY tblShiftManager.shiftStart DESC;")
    End Select
    
    btnSign.enabled = reviewStatus <> ""
    'btnComment.enabled = reviewStatus <> ""
    Dim ctl: For Each ctl In Me.Controls
        If ctl.Tag = "review" Then
            ctl.Caption = reviewStatus
            ctl.Visible = reviewStatus <> ""
        End If
    Next
End Sub

