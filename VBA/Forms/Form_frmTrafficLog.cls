VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmTrafficLog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Dim tCount As Integer
Dim fpReloadTime As Integer
Private flightID As Integer
Private isSubFormOpen As Boolean
Private shiftKeyPressed As Boolean

Public Property Let selectedFlight(f As Integer)
flightID = f
'Me.Filter = "ID = " & flightID
'Me.FilterOn = True
End Property

Public Property Get selectedFlight() As Integer
selectedFlight = flightID
'Me.Filter = "ID = " & flightID
'Me.FilterOn = True
End Property

Private Function handleMovement(ByVal movement As String)
'Called by one of the 6 flight plan movement command buttons.
'NOTE: The default InputBox here only allows one input (the time). Design a custom InputBox form for more complex inputs.
'ceta (dof,etd,ete,eta*,atd*,ata*) as date
On Error GoTo errtrap
If selectedFlight = 0 Then Exit Function

Dim db As DAO.Database: Set db = CurrentDb
Dim movetime As Variant
Dim dataObj As New MSForms.DataObject
'Dim subject As String: subject = DLookup("Callsign & ""/"" & actype & "" | "" & deppoint & "" -> "" & destination", "tblTraffic", "ID = " & selectedFlight)
Dim subject As String: subject = DLookup("Callsign & ""/"" & actype", "tblTraffic", "ID = " & selectedFlight)
Dim atlasID As Double: atlasID = Nz(DLookup("atlasID", "tblTraffic", "ID = " & selectedFlight), 0)

    Me.Filter = "ID = " & selectedFlight
    Me.FilterOn = True
    
    If Not shiftKeyPressed And (movement = "ARR" Or movement = "DEP" Or movement = "ETA") Then
        movetime = getTime4Char(InputBox("Enter the " & movement & " (Z) time for:" & vbCrLf & _
            subject, _
            "Movement", Format(LToZ(Now), "hhnn")))
    Else
        movetime = TimeValue(LToZ(Now))
    End If
            
    Select Case movement
    
    Case "ARR"
        If DLookup("status", "tbltraffic", "ID = " & selectedFlight) Like "*Enroute*" _
        Or DLookup("status", "tbltraffic", "ID = " & selectedFlight) Like "*Closed*" Then
            
            If IsNull(movetime) Then
                log "Invalid value for moveTime. (" & CStr(Nz(movetime, "<NULL>")) & ")", Me.Name & ".handleMovement", "WARN"
                'Beep
                GoTo fexit
            End If
            
            db.Execute "UPDATE tblTraffic " & _
                "SET ATA = #" & movetime & "#, " & _
                "Status = 'Closed', " & _
                "Stationed = True, " & _
                "arrDate = ztol(cETA(DOF,ETD,ETE,Null,Null,TimeValue(#" & movetime & "#))) WHERE ID = " & selectedFlight, dbFailOnError
                
        ElseIf DLookup("flightRule", "tbltraffic", "ID = " & selectedFlight) = "A" And DLookup("status", "tbltraffic", "ID = " & selectedFlight) Like "*Pending*" Then
        
            db.Execute "UPDATE tblTraffic " & _
                "SET ATA = #" & movetime & "#, " & _
                "Status = 'Enroute', " & _
                "arrDate = ztol(cETA(DOF,ETD,ETE,Null,Null,TimeValue(#" & movetime & "#))) WHERE ID = " & selectedFlight, dbFailOnError
            
        Else
            MsgBox "Cannot ARR without DEP or ETA.", vbInformation, "ARR"
        End If
            
    Case "DEP"
        If DLookup("status", "tbltraffic", "ID = " & selectedFlight) Like "*Pending*" Then
            
'            If Not shiftKeyPressed Then
'                movetime = getTime4Char(InputBox("Enter the DEP (Z) time for:" & vbCrLf & _
'                    subject, _
'                    "DEP", Format(LToZ(Now), "hhnn")))
'            Else
'                movetime = TimeValue(LToZ(Now))
'            End If
            
            If IsNull(movetime) Then
                log "Invalid value for moveTime.", Me.Name & ".handleMovement", "WARN"
                'Beep
                GoTo fexit
            End If
            
            db.Execute "UPDATE tblTraffic " & _
                "SET ATD = #" & movetime & "#, " & _
                "ETA = Null, " & _
                "Stationed = False, " & _
                "Status = 'Enroute', " & _
                "depDate = ztol(DOF + TimeValue(#" & movetime & "#)), " & _
                "arrDate = ztol(cETA(DOF,ETD,ETE,Null,TimeValue(#" & movetime & "#))) WHERE ID = " & selectedFlight, dbFailOnError

        ElseIf DLookup("flightRule", "tbltraffic", "ID = " & selectedFlight) = "A" And DLookup("status", "tbltraffic", "ID = " & selectedFlight) Like "*Enroute*" Then
        
            db.Execute "UPDATE tblTraffic " & _
                "SET ATD = #" & movetime & "#, " & _
                "Status = 'Closed', " & _
                "depDate = ztol(DOF + TimeValue(#" & movetime & "#)) " & _
                "WHERE ID = " & selectedFlight, dbFailOnError
        
        Else
            MsgBox "Cannot DEP closed/cancelled flights.", vbInformation, "DEP"
        End If
            
    Case "ETA"
    '''ID  atlasField  atlasSolution
    '14  InETA "Iif(Forms!quick_input!Destination = Dlookup(""data"",""tblsettings"",""key = 'station'""), iif(Forms!quick_input!depPoint = Dlookup(""data"",""tblsettings"",""key = 'station'""),Forms!quick_input!ETA,cETA(Forms!quick_input!DOF,Forms!quick_input!ETD,Forms!quick_input!ETE)),Null)"
        If Not DLookup("status", "tbltraffic", "ID = " & selectedFlight) Like "*Closed*" _
        And Not DLookup("status", "tbltraffic", "ID = " & selectedFlight) Like "*Cancelled*" Then
        
                    
'            If Not shiftKeyPressed Then
'                Dim gt4c As String: gt4c = InputBox("Enter the ETA (Z) time for:" & vbCrLf & _
'                    DLookup("Callsign & ""/"" & actype & "" | "" & deppoint & "" -> "" & destination", "tblTraffic", "ID = " & selectedFlight), _
'                    "ETA", Format(LToZ(Now), "hhnn"))
'
'                If Left(gt4c, 1) = "+" Then
'                    movetime = TimeValue(DateAdd("n", Mid(gt4c, InStr(1, gt4c, "+") + 1), LToZ(Now)))
'                Else
'                    movetime = getTime4Char(gt4c)
'                End If
'            Else
'                movetime = TimeValue(LToZ(Now))
'            End If
            
            If IsNull(movetime) Then
                log "Invalid value for moveTime.", Me.Name & ".handleMovement", "WARN"
                'Beep
                GoTo fexit
            End If
            
            If Not DLookup("flightRule", "tblTraffic", "ID = " & selectedFlight) = "A" Then
                db.Execute "UPDATE tblTraffic " & _
                    "SET ETA = TimeValue(#" & movetime & "#), " & _
                    "Status = 'Enroute', " & _
                    "Stationed = False, " & _
                    "arrDate = ztol(cETA(DOF,ETD,ETE,TimeValue(#" & movetime & "#))) WHERE ID = " & selectedFlight, dbFailOnError
            Else
                db.Execute "UPDATE tblTraffic " & _
                "SET ETA = TimeValue(#" & movetime & "#) " & _
                "WHERE ID = " & selectedFlight, dbFailOnError
            End If
            
        Else
            MsgBox "Cannot ETA closed/cancelled flights.", vbInformation, "ETA"
        End If
        
    Case "CLR"
        If Not DLookup("status", "tbltraffic", "ID = " & selectedFlight) Like "*Pending*" _
        And Not DLookup("status", "tbltraffic", "ID = " & selectedFlight) Like "*Cancelled*" Then
        
            If Not shiftKeyPressed Then
                If MsgBox(subject & vbCrLf & "Reset?", vbQuestion + vbYesNo, "CNL") = vbNo Then GoTo fexit
            End If
            
            db.Execute "UPDATE tblTraffic " & _
                "SET Status = 'Pending', " & _
                "ATA = Null, ATD = Null, ETA = Null, " & _
                "Stationed = (depPoint = DLookup('data','tblSettings','key = ""station""')) " & _
                "WHERE ID = " & selectedFlight, dbFailOnError
                
        End If
    
    Case "CNL"
        If Not DLookup("status", "tbltraffic", "ID = " & selectedFlight) Like "*Enroute*" _
        And Not DLookup("status", "tbltraffic", "ID = " & selectedFlight) Like "*Closed*" Then
            
            Dim q As String, a As String
            If DLookup("status", "tbltraffic", "ID = " & selectedFlight) Like "*Cancelled*" Then
                q = "Re-activate?"
                a = "Pending"
            Else
                q = "Cancel?"
                a = "Cancelled"
            End If
            
            If Not shiftKeyPressed Then
                If MsgBox(subject & vbCrLf & q, vbQuestion + vbYesNo, "CNL") = vbNo Then GoTo fexit
            End If
            
            db.Execute "UPDATE tblTraffic " & _
                "SET Status = '" & a & "', " & _
                "ATA = Null, ATD = Null " & _
                "WHERE ID = " & selectedFlight, dbFailOnError

        Else
            MsgBox "Cannot CNL without CLR.", vbInformation, "CNL"
        End If
        
    Case "RS"
        If DLookup("status", "tbltraffic", "ID = " & selectedFlight) Like "*Pending*" _
        Or DLookup("status", "tbltraffic", "ID = " & selectedFlight) Like "*Cancelled*" Then
            
            If MsgBox(subject & vbCrLf & "Remove Strip (Delete)?", vbQuestion + vbYesNo, "RS") = vbNo Then GoTo fexit
                
            If Nz(getSettings("syncAtlas"), False) And Nz(atlasID, 0) <> 0 Then
                CurrentDb.Execute "DELETE FROM atlAtlas WHERE recID = " & atlasID
            End If
            
            db.Execute "DELETE FROM tblTraffic " & _
                "WHERE ID = " & selectedFlight, dbFailOnError
        Else
            MsgBox "Cannot Remove Strip without CLR.", vbInformation, "CNL"
        End If
        
    End Select
    
fexit:
    If db.RecordsAffected <> 0 Then
        Dim movementMsg As String
        If movement <> "RS" Then
            movementMsg = subject & ": (" & movement & "-" & Mid(subject, 1, InStr(1, subject, "/") - 1) & "-" & depPoint & IIf(movement = "DEP", Format(Nz(movetime, Now), "hhnn"), "") & "-" & Destination & IIf(movement = "ARR", Format(Nz(movetime, Now), "hhnn"), "") & "-0)"
        Else
            movementMsg = subject & ": (RS)"
        End If
        
        log movementMsg, Me.Name & ".handleMovement"
        dataObj.SetText Mid(movementMsg, InStr(1, movementMsg, ":") + 2)
        dataObj.putInClipboard
        
        FormSyncUtil.syncForm Me.Name
        FormSyncUtil.syncForm "frmFlightMonitor"
'        If Nz(getSettings("syncAtlas"), False) And atlasID <> 0 And movement <> "CNL" Then
'            DoCmd.OpenForm "quick_input", , , "ID = " & selectedFlight, acFormEdit, acHidden
'            linkAtlas False, atlasID
'            DoCmd.Close acForm, "quick_input"
'        End If
        fpReload
    End If
    Set db = Nothing
    shiftKeyPressed = False
    Exit Function
errtrap:
    ErrHandler err, Error$, Me.Name & ".handleMovement"
    Resume fexit
    Resume Next
End Function

Private Sub all_pprs_Click()
Dim ti As Integer: ti = Me.TimerInterval
    Me.TimerInterval = 0
    
    DoCmd.OpenForm "frmAllPPRs", , , , , acDialog
    
    Me.TimerInterval = ti
End Sub

Private Sub btnAnnounce_Click()
    DoCmd.OpenForm "frmAnnounce"
End Sub

Private Sub bOsPPR_Click()
If IsNull(on_station) Then Exit Sub
DoCmd.OpenForm "new_ppr", , , "PPR = '" & on_station.Column(1) & "'", acFormEdit, acDialog
onPPRFormClose
End Sub

Private Sub btnAllFlights_Click()
DoCmd.OpenForm "checked_flts"
End Sub

Private Sub btnAtlasSync_Click()
If Not Nz(getSettings("syncAtlas"), False) Then Exit Sub
DoCmd.Hourglass True
 If atlasPull(Date) Then
    fpReload
    FormSyncUtil.syncForm "frmTrafficLog"
    FormSyncUtil.syncForm "frmFlightMonitor"
Else
    Beep
    log "atlasPull failed for some reason", Me.Name & ".btnAtlasSync_Click", "WARN"
End If
DoCmd.Hourglass False
End Sub

Private Sub btnNewFlight_Click()
    openFlight
End Sub

Private Sub btnNewStereo_Click()
    Dim ti As Integer: ti = Me.TimerInterval
    Me.TimerInterval = 0
    
    DoCmd.OpenForm "frmStereoFlightPlanSelect", , , , , acDialog
    fpReload 3
    
    Me.TimerInterval = ti
End Sub

Private Sub btnOnStation_Click()
    Select Case btnOnStation
        Case True
            btnOnStation.Caption = "Transient"
        Case False
            btnOnStation.Caption = "Base-Assigned"
        Case Else
            btnOnStation.Caption = "All"
    End Select
            
    osReload
End Sub

Public Function osReload()
Dim qdf As DAO.QueryDef
Set qdf = CurrentDb.QueryDefs("qOnStation")
qdf.Parameters("bos") = btnOnStation
Set on_station.Recordset = qdf.OpenRecordset
End Function

Private Sub btnPPRQuickSearch_Click()
Dim s, ibPPR As String

s = InputBox("Please enter a PPR number:", "PPR Quick Search")

DoCmd.Hourglass True
ibPPR = Nz(DLookup("ppr", "tblppr", "replace(ppr,'-','') = '" & s & "' OR ppr = '" & s & "'"))
DoCmd.Hourglass False

If ibPPR <> "" Then
    DoCmd.OpenForm "new_ppr", , , "ppr = '" & ibPPR & "'", acFormEdit, acDialog
    onPPRFormClose
Else
    MsgBox "PPR not found.", vbInformation, "PPR Quick Search"
End If
    
End Sub

Private Sub btnPPRToFltPlan_Click()
If Nz(ppr_list) = "" Then Exit Sub
Dim RS As DAO.Recordset: Set RS = CurrentDb.OpenRecordset("tblTraffic")
Dim rs2 As DAO.Recordset: Set rs2 = CurrentDb.OpenRecordset("SELECT * FROM tblPPR WHERE ID = " & ppr_list.Column(0))
Dim rs3 As DAO.Recordset: Set rs3 = CurrentDb.OpenRecordset("qInbound")
Dim initialized As Boolean
    
    'If there are inbounds, find the matching callsign
    Do While Not rs3.EOF
        initialized = (ppr_list.Column(4) = rs3!Callsign)
        If initialized Then Exit Do
        rs3.MoveNext
    Loop
        
    If Not initialized Then
    
        With RS
            .AddNew
            Dim i: For i = 0 To rs2.Fields.Count - 1
                Select Case rs2.Fields(i).Name
                Case "flightrule", "Callsign", "acType", "depPoint", "Tail", "Spot", "PPR", "arrDate", "depDate"
                    .Fields(rs2.Fields(i).Name) = Nz(rs2.Fields(i).Value)
'                Case "ETA"
'                    !DOF = DateSerial(Year(rs2.Fields("arrDate")), Month(rs2.Fields("arrDate")), Day(rs2.Fields("arrDate")))
'                    !ETA = LToZ(TimeSerial(Hour(rs2.Fields("arrDate")), Minute(rs2.Fields("arrDate")), 0))
                    
                End Select
            Next
            !flightRule = "I"
            !DOF = DateSerial(Year(rs2.Fields("arrDate")), Month(rs2.Fields("arrDate")), Day(rs2.Fields("arrDate")))
            !ETA = Format(LToZ(TimeValue(rs2.Fields("arrDate"))), "hh:nn")
            !entryDate = Now
            !AMOPS = Util.getOpInitials(Util.getUser)
            !Status = "Pending"
            !Destination = DLookup("data", "tblSettings", "key = 'station'")
            .update
        End With
        Me.inbound_aircraft.Requery
        CurrentDb.Execute "UPDATE tblSettings SET frmTrafficLogSync = Now()", dbFailOnError
    End If
    
    RS.Close
    rs2.Close
    rs3.Close
    Set RS = Nothing
    Set rs2 = Nothing
    Set rs3 = Nothing
End Sub

Private Sub btnShowCancelled_Click()
    Select Case btnShowCancelled
        Case True
            btnShowCancelled.Caption = "Hide Cancelled"
        Case False
            btnShowCancelled.Caption = "Show Cancelled"
    End Select
    pprlogReload
End Sub

Private Sub pprlogReload()
    Dim qdf As DAO.QueryDef
    Set qdf = CurrentDb.QueryDefs("qPPR")
    qdf.Parameters("showCancelled") = btnShowCancelled
    Set ppr_list.Recordset = qdf.OpenRecordset
    ppr_list.Requery
End Sub

Private Sub btnViewNotice_Click()
DoCmd.OpenForm "frmAnnounce"
End Sub

Private Sub dep_arr_lcl_Click()
On Error GoTo check_out_Click_err
    DoCmd.OpenForm "check_in_out", , , "ID=" & local_aircraft.Column(0), acFormEdit
check_out_Click_err:
    Exit Sub
End Sub

Private Sub bCancelPPR_Click()
On Error GoTo bCancelPPR_Click_err
Dim db As DAO.Database
Dim RS As DAO.Recordset
Dim rstSP As DAO.Recordset
Dim qdf As DAO.QueryDef
Dim clm As Integer
Dim sts As String
Set db = CurrentDb

    If IsNull(ppr_list.Column(0)) Then Exit Sub
        
    clm = ppr_list.Column(0)
    sts = DLookup("Status", "tblPPR", "ID=" & clm)
    Set RS = db.OpenRecordset("SELECT * FROM tblPPR WHERE tblPPR.ID = " & clm)
    If Not sts = "Cancelled" Then
        If MsgBox("Cancel PPR?", vbYesNo, "Confirm") = vbYes Then
            RS.edit
            RS!Status = "Cancelled"
            RS!approval = 0
            RS.update
            If DLookup("spAccess", "tblUserAuth", "username = '" & Util.getUser & "'") And Not IsNull(RS!spID) Then
                Set qdf = db.QueryDefs("qMissionTrackerByID")
                qdf.Parameters("mtbid") = RS!spID
                Set rstSP = qdf.OpenRecordset()
                With rstSP
                    .edit
                    ![Mission/Event Type] = "Cancelled"
                    .update
                End With
            End If
        Else
            Exit Sub
        End If
    Else
        If MsgBox("Re-activate?", vbYesNo, "Confirm") = vbYes Then
            If Not RS.EOF Then
                RS.edit
                RS!Status = IIf(RS!approach, "Approach", "Approved")
                RS!approval = 1
                RS.update
                If DLookup("spAccess", "tblUserAuth", "username='" & Util.getUser & "'") And Not IsNull(RS!spID) Then
                    Set qdf = db.QueryDefs("qMissionTrackerByID")
                    qdf.Parameters("mtbid") = RS!spID
                    Set rstSP = qdf.OpenRecordset()
                    With rstSP
                        .edit
                        ![Mission/Event Type] = "PPR"
                        .update
                    End With
                    rstSP.Close
                    Set rstSP = Nothing
                End If
            End If
        Else
            Exit Sub
        End If
    End If

'    rs.close
'    Set rs = Nothing
    
    syncTrafficLog RS!ID, "tblPPR", False
    ppr_list.Requery
    
bCancelPPR_Click_Exit:
    Exit Sub
    
bCancelPPR_Click_err:
    MsgBox Error$
    Resume Next
End Sub

Private Sub bNewPPR_Click()
On Error GoTo errtrap
    Dim ti As Integer: ti = Me.TimerInterval
    Me.TimerInterval = 0
    
    isSubFormOpen = True
    DoCmd.OpenForm "new_ppr", , , , acFormAdd, acDialog
    onPPRFormClose
    
sexit:
    Me.TimerInterval = ti
    Exit Sub
errtrap:
    isSubFormOpen = False
    GoTo sexit
End Sub

Private Sub checked_Click()
    DoCmd.OpenForm "checked_flts"
End Sub

Private Sub dep_arr_in_Click()
On Error GoTo check_out_Click_err
    DoCmd.OpenForm "check_in_out", , , "ID=" & inbound_aircraft.Column(0), acFormEdit
check_out_Click_err:
    Exit Sub
End Sub

Private Sub dep_out_Click()
On Error GoTo check_out_Click_err
    DoCmd.OpenForm "check_in_out", , , "ID=" & outbound_aircraft.Column(0), acFormEdit
check_out_Click_err:
    Exit Sub
End Sub

Private Sub Detail_Click()
    Dim ctl: For Each ctl In Controls
        If TypeOf ctl Is ListBox Then ctl.Value = ""
    Next
    selectedFlight = 0
End Sub

Private Sub feedback_Click()
    DoCmd.OpenForm "frmFeedback", , , , acFormAdd
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
    Case vbKeyShift
        shiftKeyPressed = True
    
    End Select
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
    Case vbKeyShift
        shiftKeyPressed = False
    End Select
End Sub

Private Sub Form_Load()
    
    Dim ctl: For Each ctl In Controls
        If ctl.Tag Like "*kiosk*" Then
        End If
    Next
    
    If DLookup("data", "tblsettings", "key = 'syncAtlasOnStart'") Then TrafficUtil.atlasPull Date
    FormSyncUtil.syncForm Me.Name, True
    
    Dim qdfAnnounce As DAO.QueryDef: Set qdfAnnounce = CurrentDb.QueryDefs("qAnnounce")
    qdfAnnounce.Parameters("cd1") = Date
    qdfAnnounce.Parameters("cd2") = Date
    Set lstAnnounce.Recordset = qdfAnnounce.OpenRecordset
    
    Me.TimerInterval = DLookup("data", "tblsettings", "key = 'globalSyncLength'")
    
    osReload
    pprlogReload
    fpReload
End Sub

Private Sub Form_Timer()
On Error Resume Next
If isSubFormOpen Then Exit Sub

    If Not FormSyncUtil.isFormSynced(Me.Name) Or fpReloadTime >= DLookup("data", "tblsettings", "key = 'globalSyncInterval'") Then
        If Nz(getSettings("syncAtlas"), False) Then TrafficUtil.atlasPull Date
        
        osReload
        pprlogReload
        fpReload
        
        lstAnnounce.Requery
        
        CurrentDb.Execute "qTrafficFlightStatusUpdate"
        
        FormSyncUtil.syncForm Me.Name, True
        fpReloadTime = 0
    End If
    
'    Dim qdf As DAO.QueryDef: Set qdf = CurrentDb.QueryDefs("qLocal")
'    Dim rs As DAO.Recordset
'    With qdf
'        .Parameters("cd1") = Date
'        .Parameters("cd2") = Date
'        Set rs = .OpenRecordset
'    End With
'
'    With rs: Do While Not .EOF
'        Select Case !Status
'        Case "Pending"
'            If !ETD < DateAdd("h", -1, Now) Then
'                Me.SetFocus
'                local_aircraft.BackColor = vbYellow
'                local_aircraft.ForeColor = vbRed
'                lblLateLcl.Visible = True
'            End If
'
'        Case "*Enroute*"
'            If !ETA < DateAdd("n", -30, Now) Then
'                Me.SetFocus
'                local_aircraft.BackColor = vbYellow
'                local_aircraft.ForeColor = vbRed
'                lblLateLcl.Visible = True
'            End If
'
'    End Select: Loop: End With
        
            

fpReloadTime = fpReloadTime + 1
End Sub

Private Sub inbound_aircraft_Click()
    selectedFlight = inbound_aircraft
End Sub

Private Sub inbound_aircraft_DblClick(cancel As Integer)
    If Not IsNull(Me.inbound_aircraft.Column(0)) Then
'        DoCmd.OpenForm "quick_input", , , "ID=" & Me.inbound_aircraft.Column(0), acFormEdit
        openFlight Me.inbound_aircraft.Column(0)
    End If
End Sub

Private Sub inbound_aircraft_GotFocus()
    outbound_aircraft = ""
    local_aircraft = ""
    on_station = ""
    ppr_list = ""
End Sub

Private Sub inCancel_Click()
    If inbound_aircraft = "" Then Exit Sub
    If cnlFlight(CurrentDb.OpenRecordset("SELECT * FROM tblTraffic WHERE tblTraffic.ID = " & Nz(inbound_aircraft.Column(0)) & ";")) Then inbound_aircraft.Requery
End Sub

Private Sub lclCancel_Click()
    If cnlFlight(CurrentDb.OpenRecordset("SELECT * FROM tblTraffic WHERE tblTraffic.ID = " & Nz(local_aircraft.Column(0)) & ";")) Then local_aircraft.Requery
End Sub

Private Sub local_aircraft_Click()
    selectedFlight = local_aircraft
End Sub

Private Sub local_aircraft_DblClick(cancel As Integer)
    If Not IsNull(Me.local_aircraft.Column(0)) Then
        'DoCmd.OpenForm "quick_input", , , "ID=" & Me.local_aircraft.Column(0), acFormEdit
        openFlight Me.local_aircraft.Column(0)
    End If
End Sub

Private Sub local_aircraft_GotFocus()
    inbound_aircraft = ""
    outbound_aircraft = ""
    on_station = ""
    ppr_list = ""
End Sub

Private Sub menu_Click()
    DoCmd.Close acForm, "Traffic Log", acSaveNo
    DoCmd.OpenForm "frmMain"
End Sub

Private Sub new_flt_Click()
    'DoCmd.OpenForm "quick_input", , , , acFormAdd, acDialog
    'Forms!quick_input.cancel_flt.Visible = False
    
    openFlight
End Sub

Private Sub openFlight(Optional ID As Integer, Optional onStation As Boolean)
    Dim ti As Integer: ti = Me.TimerInterval
    Me.TimerInterval = 0
    
    DoCmd.OpenForm "quick_input", , , IIf(ID <> 0, "ID = " & ID, ""), IIf(ID <> 0, acFormEdit, acFormAdd), acDialog, IIf(onStation, 4, Null)
    If CurrentProject.AllForms("quick_input").IsLoaded Then
        fpReload
        DoCmd.Close acForm, "quick_input"
    End If
    
    Me.TimerInterval = ti
End Sub

Private Sub fpReload(Optional FPType As Integer)
    Select Case FPType
    Case 0 ' Nothing received; update all windows
        Dim ctl: For Each ctl In Controls
            If TypeOf ctl Is ListBox Then
                If ctl.Tag Like "*fp*" Then ctl.Requery
            End If
        Next
    
    Case 1 'Update inbound only
        inbound_aircraft.Requery
    Case 2 'Update outbound only
        outbound_aircraft.Requery
    Case 3 'Update local only
        local_aircraft.Requery
    End Select
    
    osReload
End Sub

Private Sub on_station_DblClick(cancel As Integer)
    openFlight on_station.Column(0), True
    osReload
End Sub

Private Sub on_station_GotFocus()
    inbound_aircraft = ""
    outbound_aircraft = ""
    local_aircraft = ""
    ppr_list = ""
End Sub

Private Sub os_add_Click()
    DoCmd.OpenForm "frmAddAcft", , , , acFormAdd
End Sub

Private Sub os_del_Click()
Dim db As DAO.Database
Dim RS As DAO.Recordset
Dim clm As Integer

Set db = CurrentDb

    If IsNull(on_station.Column(0)) Then Exit Sub
    clm = on_station.Column(0)
    If MsgBox("Remove aircraft from station?", vbYesNo, "Confirm") = vbYes Then
'        Set qdf = db.QueryDefs("qOnStation")
'        qdf.Parameters("bos") = btnOnStation
'        Set rs = qdf.OpenRecordset
'        Do Until rs.EOF
'            If rs!ID = clm Then
'                rs.edit
'                rs!Stationed = False
'                rs.update
'            End If
'            rs.MoveNext
'        Loop
'        rs.Close
'        Set rs = Nothing
        CurrentDb.Execute "UPDATE tblTraffic SET stationed = false WHERE ID = " & clm
        osReload
    End If
End Sub

Private Sub outbound_aircraft_Click()
    selectedFlight = outbound_aircraft
End Sub

Private Sub outbound_aircraft_DblClick(cancel As Integer)
On Error GoTo check_out_Click_err
    If Not IsNull(Me.outbound_aircraft.Column(0)) Then
        'DoCmd.OpenForm "quick_input", , , "ID=" & Me.outbound_aircraft.Column(0), acFormEdit
        openFlight Me.outbound_aircraft.Column(0)
    End If
    
check_out_Click_err:
    Exit Sub
End Sub

Private Sub outCancel_Click()
    If cnlFlight(CurrentDb.OpenRecordset("SELECT * FROM tblTraffic WHERE tblTraffic.ID = " & Nz(outbound_aircraft.Column(0)) & ";")) Then outbound_aircraft.Requery
End Sub

Private Sub ppr_list_DblClick(cancel As Integer)

    If Not IsNull(Me.ppr_list.Column(0)) Then
        DoCmd.OpenForm "new_ppr", , , "ID=" & ppr_list.Column(0), acFormEdit, acDialog
        onPPRFormClose
    End If
End Sub

Private Sub onPPRFormClose()
If Not CurrentProject.AllForms("new_ppr").IsLoaded Then Exit Sub

    If Forms!new_ppr.saveResult Then
        pprlogReload
        osReload
    End If
    DoCmd.Close acForm, "new_ppr"
End Sub

Private Sub outbound_aircraft_GotFocus()
    inbound_aircraft = ""
    local_aircraft = ""
    on_station = ""
    ppr_list = ""
End Sub

Private Sub ppr_list_GotFocus()
    inbound_aircraft = ""
    outbound_aircraft = ""
    local_aircraft = ""
    on_station = ""
End Sub


Private Sub pprSync_Click()
On Error GoTo errtrap
Dim RS As DAO.Recordset
Dim rsSP As DAO.Recordset
Dim rsPPR As DAO.Recordset
Dim s() As String
Dim idx As Integer
Set RS = ppr_list.Recordset
RS.Sort = "spID"
Set rsPPR = RS.OpenRecordset
tCount = 0

    Dim i: For i = 1 To ppr_list.ListCount - 1
        If ppr_list.Column(1, i) <> "" Then
            ReDim Preserve s(0 To idx) As String
            s(idx) = ppr_list.Column(1, i)
            idx = idx + 1
        End If
    Next
    
   
    DoEvents
    
    Set rsSP = CurrentDb.OpenRecordset("SELECT * FROM spPPRLog WHERE ID = " & join(s, " OR ID = ") & " ORDER BY ID")
    With rsSP: Do While Not .EOF
        With rsPPR: Do While Not .EOF
            If !spID = rsSP!ID Then
                .edit
                !Spot = rsSP![Parking Spot/Location]
                .update
                rsSP.MoveNext
            End If
            .MoveNext
        Loop: End With
    Loop: End With
    ppr_list.Requery
    
    DoEvents
    syncTrafficLog
    
sexit:
    Exit Sub
errtrap:
    ErrHandler err, Error$, Name
    Resume Next
End Sub
