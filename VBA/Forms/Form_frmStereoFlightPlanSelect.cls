VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmStereoFlightPlanSelect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
Private bClose As Boolean
Private seq As Integer

Public Function isValid() As Boolean
isValid = True

    Dim ctl: For Each ctl In Me.Controls
        If ctl.Tag Like "*valid*" Then
            If Nz(ctl.Value) = "" Then
            
                ctl.BorderColor = RGB(255, 0, 0)
                isValid = False
                
            Else
                ctl.BorderColor = RGB(166, 166, 166)
                
            End If
        End If
    Next ctl
    
    If Not isValid Then Beep
    lblAISR.Visible = isValid
    AISR.Visible = isValid
End Function

Private Sub AISR_Click()
If Not isValid Then Exit Sub
Dim dataObj As New MSForms.DataObject
    With dataObj

        .SetText AISR
        .putInClipboard
        log "Copied AISR string: " & AISR, Me.Name & ".AISR_Click"
        Beep
    End With
    Set dataObj = Nothing
End Sub

Private Sub btnSave_Click()
On Error GoTo errtrap
If Not isValid Then Exit Sub
DoEvents
Dim Stereo As Integer: Stereo = Nz(lstStereo.Column(0), 0): If Stereo = 0 Then Exit Sub
Dim RS As DAO.Recordset: Set RS = CurrentDb.OpenRecordset("SELECT * FROM tblStereoFlightPlan WHERE ID = " & Stereo)

    lblLoading.Visible = True
    Dim ctl: For Each ctl In Me.Controls
        If Not TypeOf ctl Is Label Then
            ctl.Enabled = False
        End If
    Next ctl
            
    
    With RS
        If .EOF Then Exit Sub
        
        If DCount("ID", "tblTraffic", "ID = " & Me.ID) > 1 Then Me.ID = DMax("ID", "tblTraffic") + 1
        
        Me.flightRule = "S"
        Me.Status = "Pending"
        Me.depPoint = DLookup("data", "tblSettings", "key = 'station'")
        Me.Destination = DLookup("data", "tblSettings", "key = 'station'")
        
        Me.Callsign = UCase(Me.Callsign)
        If IsNull(Me.acType) Then Me.acType = !acType
        Me.ETE = !ETE
        
        Me.arrDate = ZToL(cETA(DOF, ETD, ETE))
        Me.depDate = ZToL(DOF + ETD)
        
        'if datediff("d",now,ZToL(cETA(DOF, ETD, ETE))) <> 0
        
        Me.stereoPlan = lstStereo.Column(1)
        
        Me.rwy = DLast("rwy", "tblshiftmanager")
        
        Me.Remarks = Nz(openargs) & getOpInitials(getUser) & vbCrLf & Nz(Me.Remarks)
        DoEvents
        'Me.altitude = !altitude
        Call AISR_Click
        
        
        Me.AMOPS = getOpInitials(Util.getUser)
    End With
    
    bClose = True
    If Dirty Then Dirty = False
    
'    If btnAtlas Then
'        DoCmd.OpenForm "quick_input", , , "ID = " & ID, acFormEdit, acHidden
'        Me.atlasID = linkAtlas(NewRecord, Nz(atlasID, 0))
'        DoCmd.Close acForm, "quick_input"
'    End If
    
cleanup:
    lblLoading.Visible = False
    For Each ctl In Me.Controls
        If Not TypeOf ctl Is Label Then
            ctl.Enabled = True
        End If
    Next ctl
    
sexit:
    RS.Close
    Set RS = Nothing
    FormSyncUtil.syncForm "frmTrafficLog"
    FormSyncUtil.syncForm "frmFlightMonitor"
    Me.Visible = False
    Exit Sub
errtrap:
    If err = 3022 Then Resume Next
    ErrHandler err, Error$, Me.Name & ".btnSave_Click"
    Resume sexit
    Resume Next
End Sub

Private Sub btnSettings_Click()
DoCmd.OpenForm "frmStereoFlightPlan", , , , , acDialog
lstStereo.Requery
End Sub

Private Sub ctlETD_Exit(cancel As Integer)
    If Nz(ctlETD) = "" Then Exit Sub
    If IsDate(ctlETD) Then ctlETD = Format(ctlETD, "hhnn")
    If IsNumeric(ctlETD) And Len(ctlETD) = 3 Then ctlETD = 0 & ctlETD
    
    ctlETD.BorderColor = RGB(166, 166, 166)
    ETD = Nz(getTime4Char(ctlETD))
    ctlETD = Left(Format(ETD, "hhnn"), 2) & ":" & Right(Format(ETD, "hhnn"), 2)
    
    If Not IsDate(ETD) Then
        'cancel = True
        ctlETD = ""
        ctlETD.BorderColor = RGB(255, 0, 0)
        ctlETD.SetFocus
        Beep
    End If
End Sub

Private Sub Form_Current()
    bClose = False
    If Not IsNull(ETD) Then ctlETD = Hour(ETD) & ":" & Format(Minute(ETD), "00")
End Sub

Private Sub Form_BeforeUpdate(cancel As Integer)
    If Not bClose Then
        'cancel = True
        Undo
        Exit Sub
    End If
End Sub

Private Sub lstStereo_BeforeUpdate(cancel As Integer)
    If IsNull(acType) Then acType = lstStereo.Column(2)
    
    If Not isValid Then
        'cancel = True
    Else
        lblAISR.Visible = isValid
        AISR.Visible = isValid
    End If
End Sub

Private Sub lstStereo_DblClick(cancel As Integer)
MsgBox DLookup("""["" & stereotag & ""] "" & iif(number>1,number & ""/"","""") & actype & ""/"" & equipment & "" "" & speed & "" "" & altitude & "" "" & route", _
        "tblStereoFlightPlan", _
        "ID = " & lstStereo), _
    vbInformation, "Stereo Plan"
End Sub
