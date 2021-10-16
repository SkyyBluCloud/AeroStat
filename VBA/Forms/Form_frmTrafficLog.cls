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
End Property

Public Property Get selectedFlight() As Integer
selectedFlight = flightID
End Property

Private Function handleMovement(ByVal movement As String)
'Called by one of the 6 flight plan movement command buttons.
'NOTE: The default InputBox here only allows one input (the time). Design a custom InputBox form for more complex inputs.
'ceta (dof,etd,ete,eta*,atd*,ata*) as date
On Error GoTo errtrap
If selectedFlight = 0 Then Exit Function

Dim db As DAO.Database: Set db = CurrentDb
Dim moveTime As Variant
'Dim subject As String: subject = DLookup("Callsign & ""/"" & actype & "" | "" & deppoint & "" -> "" & destination", "tblTraffic", "ID = " & selectedFlight)
Dim subject As String: subject = DLookup("Callsign & ""/"" & actype", "tblTraffic", "ID = " & selectedFlight)
Dim atlasID As Double: atlasID = Nz(DLookup("atlasID", "tblTraffic", "ID = " & selectedFlight), 0)

    Select Case movement
    
    Case "ARR"
        If DLookup("status", "tbltraffic", "ID = " & selectedFlight) Like "*Enroute*" _
        Or DLookup("status", "tbltraffic", "ID = " & selectedFlight) Like "*Closed*" Then
            
            If Not shiftKeyPressed Then
                moveTime = getTime4Char(InputBox("Enter the ARR (Z) time for:" & vbCrLf & _
                    subject, _
                    "ARR", Format(LToZ(Now), "hhnn")))
            Else
                moveTime = TimeValue(LToZ(Now))
            End If
            
            If IsNull(moveTime) Then
                log "Invalid value for moveTime. (" & CStr(Nz(moveTime, "<NULL>")) & ")", Me.Name & ".handleMovement", "WARN"
                'Beep
                GoTo fexit
            End If
            
            db.Execute "UPDATE tblTraffic " & _
                "SET ATA = #" & moveTime & "#, " & _
                "Status = 'Closed', " & _
                "Stationed = True, " & _
                "arrDate = ztol(cETA(DOF,ETD,ETE,Null,Null,TimeValue(#" & moveTime & "#))) WHERE ID = " & selectedFlight, dbFailOnError
                
        Else
            MsgBox "Cannot ARR without DEP or ETA.", vbInformation, "ARR"
        End If
            
    Case "DEP"
        If Not DLookup("status", "tbltraffic", "ID = " & selectedFlight) Like "*Closed*" _
        And Not DLookup("status", "tbltraffic", "ID = " & selectedFlight) Like "*Cancelled*" Then
            
            If Not shiftKeyPressed Then
                moveTime = getTime4Char(InputBox("Enter the DEP (Z) time for:" & vbCrLf & _
                    subject, _
                    "DEP", Format(LToZ(Now), "hhnn")))
            Else
                moveTime = TimeValue(LToZ(Now))
            End If
            
            If IsNull(moveTime) Then
                log "Invalid value for moveTime.", Me.Name & ".handleMovement", "WARN"
                'Beep
                GoTo fexit
            End If
            
            db.Execute "UPDATE tblTraffic " & _
                "SET ATD = #" & moveTime & "#, " & _
                "ETA = Null, " & _
                "Stationed = False, " & _
                "Status = 'Enroute', " & _
                "depDate = ztol(DOF + TimeValue(#" & moveTime & "#)), " & _
                "arrDate = ztol(cETA(DOF,ETD,ETE,Null,TimeValue(#" & moveTime & "#))) WHERE ID = " & selectedFlight, dbFailOnError

        Else
            MsgBox "Cannot DEP closed/cancelled flights.", vbInformation, "DEP"
        End If
            
    Case "ETA"
        If Not DLookup("status", "tbltraffic", "ID = " & selectedFlight) Like "*Closed*" _
        And Not DLookup("status", "tbltraffic", "ID = " & selectedFlight) Like "*Cancelled*" Then
            
            
                    
                    
                    
            If Not shiftKeyPressed Then
                Dim gt4c As String: gt4c = InputBox("Enter the ETA (Z) time for:" & vbCrLf & _
                    DLookup("Callsign & ""/"" & actype & "" | "" & deppoint & "" -> "" & destination", "tblTraffic", "ID = " & selectedFlight), _
                    "ETA", Format(LToZ(Now), "hhnn"))
                
                If Left(gt4c, 1) = "+" Then
                    moveTime = DateAdd("n", Mid(gt4c, 2), LToZ(Now))
                Else
                    moveTime = getTime4Char(gt4c)
                End If
            Else
                moveTime = TimeValue(LToZ(Now))
            End If
            
            If IsNull(moveTime) Then
                log "Invalid value for moveTime.", Me.Name & ".handleMovement", "WARN"
                'Beep
                GoTo fexit
            End If
            
            db.Execute "UPDATE tblTraffic " & _
                "SET ETA = TimeValue(#" & moveTime & "#), " & _
                "Status = 'Enroute', " & _
                "Stationed = True, " & _
                "arrDate = ztol(cETA(DOF,ETD,ETE,TimeValue(#" & moveTime & "#))) WHERE ID = " & selectedFlight, dbFailOnError

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
                "ATA = Null, ATD = Null, " & _
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
            
            db.Execute "DELETE FROM tblTraffic " & _
                "WHERE ID = " & selectedFlight, dbFailOnError
                
            If atlasID <> 0 Then
                CurrentDb.Execute "DELETE FROM atlAtlas WHERE recID = " & atlasID
                atlasID = 0
            End If
            
        Else
            MsgBox "Cannot Remove Strip without CLR.", vbInformation, "CNL"
        End If
        
    End Select
    
fexit:
    If db.RecordsAffected <> 0 Then
        log subject & ": (" & movement & "-" & Mid(subject, 1, InStr(1, subject, "/") - 1) & "-" & depPoint & IIf(movement = "DEP", Format(Nz(moveTime, Now), "hhnn"), "") & "-" & Destination & IIf(movement = "ARR", Format(Nz(moveTime, Now), "hhnn"), "") & ")", Me.Name & ".handleMovement"
        FormSyncUtil.syncForm Me.Name
        FormSyncUtil.syncForm "frmFlightMonitor"
        If atlasID <> 0 And movement <> "CNL" Then
            DoCmd.OpenForm "quick_input", , , "ID = " & selectedFlight, acFormEdit, acHidden
            linkAtlas False, atlasID
            DoCmd.Close acForm, "quick_input"
        End If
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
    DoCmd.OpenForm "frmAllPPRs"
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

Private Sub btnNewFlight_Click()
    openFlight
End Sub

Private Sub btnNewStereo_Click()
    DoCmd.OpenForm "frmStereoFlightPlanSelect", , , , , acDialog
    fpReload 3
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
            !ETA = Format(LToZ(TimeSerial(Hour(rs2.Fields("arrDate")), Minute(rs2.Fields("arrDate")), 0)), "hh:nn")
            !entryDate = Now
            !AMOPS = Util.getOpInitials(getUSN)
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
            If DLookup("spAccess", "tblUserAuth", "username = '" & Environ$("username") & "'") And Not IsNull(RS!spID) Then
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
                If DLookup("spAccess", "tblUserAuth", "username='" & Environ$("username") & "'") And Not IsNull(RS!spID) Then
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
    isSubFormOpen = True
    DoCmd.OpenForm "new_ppr", , , , acFormAdd, acDialog
    onPPRFormClose
errtrap:
    isSubFormOpen = False
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
    osReload
    pprlogReload
    FormSyncUtil.syncForm Me.Name, True
End Sub

Private Sub Form_Timer()
On Error Resume Next
If isSubFormOpen Then Exit Sub

    
    
    If Not FormSyncUtil.isFormSynced(Me.Name) Or fpReloadTime >= 12 Then
        osReload
        pprlogReload
        fpReload
        
        CurrentDb.Execute "qTrafficFlightStatusUpdate"
        
        FormSyncUtil.syncForm Me.Name, True
        fpReloadTime = 0
    End If

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
    DoCmd.OpenForm "quick_input", , , IIf(ID <> 0, "ID = " & ID, ""), IIf(ID <> 0, acFormEdit, acFormAdd), acDialog, IIf(onStation, 4, Null)
    If CurrentProject.AllForms("quick_input").IsLoaded Then
        fpReload Forms!quick_input.FPType
        DoCmd.Close acForm, "quick_input"
    End If
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
    
    lblWait.Caption = "Getting parking assignment from SP..."
    lblWait.Visible = True
    DoEvents
    
    Set rsSP = CurrentDb.OpenRecordset("SELECT * FROM [Mission Tracker] WHERE ID = " & join(s, " OR ID = ") & " ORDER BY ID")
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
    
    lblWait.Visible = False
    DoEvents
    syncTrafficLog
    
sexit:
    Exit Sub
errtrap:
    ErrHandler err, Error$, Name
    Resume Next
End Sub
