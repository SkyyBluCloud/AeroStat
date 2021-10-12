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
        Dim qdf As DAO.QueryDef: Set qdf = CurrentDb.QueryDefs("qTrafficStereo")
        Dim RS As DAO.Recordset
        With qdf
            .Parameters("varDate") = Date
            Set RS = .OpenRecordset
            With RS
                If Not .EOF Then
                    .MoveLast
                    .MoveFirst
                End If
                seq = .RecordCount + 1
                Me.Tail = 100 + seq
            End With
        End With
        
        Dim timestamp As String: timestamp = Right(DLookup("data", "tblSettings", "key = 'station'"), 3) & _
                                            Format(LToZ(Now), "hhnn") & _
                                            Format(seq, "000")
        .SetText timestamp & " " & AISR
        .putInClipboard
        log "Copied AISR string: " & AISR, Me.Name & ".AISR_Click"
        Beep
    End With
    Set dataObj = Nothing
    RS.Close
    Set RS = Nothing
End Sub

Private Sub btnSave_Click()
On Error GoTo errtrap
If Not isValid Then Exit Sub
DoEvents
Dim stereo As Integer: stereo = Nz(lstStereo.Column(0), 0): If stereo = 0 Then Exit Sub
Dim RS As DAO.Recordset: Set RS = CurrentDb.OpenRecordset("SELECT * FROM tblStereoFlightPlan WHERE ID = " & stereo)

    With RS
        If .EOF Then Exit Sub
        
        Me.flightRule = "S"
        Me.Status = "Pending"
        Me.depPoint = DLookup("data", "tblSettings", "key = 'station'")
        Me.Destination = DLookup("data", "tblSettings", "key = 'station'")
        
        Me.Callsign = UCase(Me.Callsign)
        Me.acType = !acType
        Me.ETE = !ETE
        
        Me.arrDate = ZToL(cETA(DOF, ETD, ETE))
        Me.depDate = ZToL(DOF + ETD)
        Me.stereoPlan = lstStereo.Column(1)
        
        Me.rwy = DLast("rwy", "tblshiftmanager")
        
        'Me.Remarks = lstStereo.Column(1)
        
        'Me.altitude = !altitude
        Call AISR_Click
        
        
        Me.AMOPS = getOpInitials(getUSN)
    End With
    
    bClose = True
    If Dirty Then Dirty = False
    
    If btnAtlas Then
        DoCmd.OpenForm "quick_input", , , "ID = " & ID, acFormEdit, acHidden
        Me.atlasID = linkAtlas(NewRecord, Nz(atlasID, 0))
        DoCmd.Close acForm, "quick_input"
    End If
    
    
    
sexit:
    RS.Close
    Set RS = Nothing
    FormSyncUtil.syncForm "frmTrafficLog"
    FormSyncUtil.syncForm "frmFlightMonitor"
    DoCmd.Close acForm, Me.Name
    Exit Sub
errtrap:
    ErrHandler err, Error$, Me.Name & ".btnSave_Click"
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
    acType = lstStereo.Column(2)
    
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
