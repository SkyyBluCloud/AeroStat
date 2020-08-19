VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_Traffic Log"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Dim tCount As Integer
Private isSubFormOpen As Boolean

Private Sub alertClear_Click()
Set RS = CurrentDb.OpenRecordset("SELECT * FROM tblUserAuth WHERE username = '" & Environ$("username") & "'")
With RS
    .edit
    !frmtrafficlogalert = alert.ItemData(alert.ListCount - 1)
    .update
    refreshAlert
End With
End Sub

Private Sub all_pprs_Click()
    DoCmd.OpenForm "frmAllPPRs"
End Sub

Private Sub announcement_GotFocus()
    menu.SetFocus
End Sub

Private Sub btnAnnounce_Click()
    DoCmd.OpenForm "frmAnnounce"
End Sub

Private Sub bOsPPR_Click()
If IsNull(on_station) Then Exit Sub
DoCmd.OpenForm "new_ppr", , , "PPR = '" & on_station.Column(1) & "'", acFormEdit, acDialog
onPPRFormClose
End Sub

Private Sub btnAddNotice_Click()
    DoCmd.OpenForm "frmAnnounceSet", , , , acFormAdd, acDialog
    lstNotice.Requery
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

    Do While Not rs3.EOF
        initialized = (ppr_list.Column(4) = rs3!Callsign)
        If initialized Then Exit Do
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
            !AMOPS = Util.getOpInitials
            !Status = "Pending"
            !Destination = DLookup("station", "tblSettings")
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
    For Each ctl In Controls
        If TypeOf ctl Is ListBox Then ctl.Value = ""
    Next
End Sub

Private Sub feedback_Click()
    DoCmd.OpenForm "frmFeedback", , , , acFormAdd
End Sub

'Private Sub Form_GotFocus()
'For Each ctl In Controls
'        If TypeOf ctl Is ListBox Then ctl.Requery
'    Next
'    chkAnnounce
'End Sub

Private Sub Form_Load()
Dim update As DAO.Recordset
Set update = CurrentDb.OpenRecordset("lclver")

    'If Not DLookup("spAccess", "tblUserAuth", "username='" & Environ$("username") & "'") Then sperr.Visible = True
    osReload
    pprlogReload
    refreshAlert
    With update
        .edit
        !frmTrafficLogSync = DLookup("frmTrafficLogSync", "tblsettings")
        .update
    End With
    
'    For Each ctl In Me.Controls
'        If ctl.Tag Like "*notice*" Then
'            ctl.Visible = lstNotice.ListCount <> 0
'        End If
'    Next
End Sub

Private Sub Form_Timer()
On Error Resume Next
If isSubFormOpen Then Exit Sub
Dim update As DAO.Recordset
Set update = CurrentDb.OpenRecordset("lclver")

With update
    lclClock.Requery
    If !frmTrafficLogSync <> DLookup("frmTrafficLogSync", "tblSettings") Then
        .edit
        !frmTrafficLogSync = DLookup("frmTrafficLogSync", "tblSettings")
        .update
        
        CurrentDb.Execute "qTrafficFlightStatusUpdate"
'        For Each ctl In Controls
'            If TypeOf ctl Is ListBox Then ctl.Requery
'        Next
        
        refreshAlert
        osReload
        pprlogReload
        fpReload
    End If
    .Close
End With
Set update = Nothing
End Sub

Public Function refreshAlert(Optional All As Boolean)
'Dim ud As String
'ud = Left(btnUpdate.Caption, 4)
'    new_flt.SetFocus
'    If All Then
'        alert.RowSource = "SELECT tblTrafficLogAlert.id, tblTrafficLogAlert.timestamp, tblTrafficLogAlert.msg FROM tblTrafficLogAlert ORDER BY timestamp DESC;"
'    Else
'        alert.RowSource = "SELECT tblTrafficLogAlert.id, tblTrafficLogAlert.timestamp, tblTrafficLogAlert.msg FROM tblTrafficLogAlert WHERE (((tblTrafficLogAlert.id)>" & DLookup("frmtrafficlogalert", "tbluserauth", "username = '" & Environ$("username") & "'") & ")) ORDER BY timestamp DESC;"
'    End If
'    alert.Visible = (ud = "Show") And ((alert.ListCount >= 1) And alert.RowSource <> "")
'    ud = IIf(Not alert.Visible, "Show", "Hide")
'
'    btnUpdate.Caption = Replace(btnUpdate.Caption, IIf(ud = "Hide", "Show", "Hide"), IIf(ud = "Hide", "Hide", "Show"))
    
'    Set rs = CurrentDb.OpenRecordset("SELECT * FROM tblUserAuth WHERE username = '" & Environ$("username") & "'")
'    With rs
'    If Not IsNull(alert.ItemData(0)) Then
'        .edit
'        !frmtrafficlogalert = alert.ItemData(0)
'        .update
'    End If
    
    
End Function

Private Sub btnUpdate_Click()
Dim RS As DAO.Recordset
Set RS = CurrentDb.OpenRecordset("SELECT * FROM tblUserAuth WHERE username = '" & Environ$("username") & "'")
'Set rs = CurrentDb.OpenRecordset("lclver")
'    With rs
'        .edit
'        !frmTrafficLogSync = DLookup("frmTrafficLogSync", "tblSettings")
'        .update
'    End With
With RS
    If alert.Visible Then
        .edit
        !frmtrafficlogalert = DMax("ID", "tblTrafficLogAlert")
        .update
    End If
    refreshAlert (!frmtrafficlogalert = DMax("ID", "tblTrafficLogAlert"))
End With
End Sub

'Private Function chkAnnounce()
'Dim rs As DAO.Recordset
'Dim db As DAO.Database
'Dim anc As String
'Dim printAnc As Boolean
'Set db = CurrentDb
'Set rs = db.OpenRecordset("SELECT * FROM tblAnnounce WHERE Format(Date(),'mm/dd/yy') Between Format(startDate,'mm/dd/yy') And Format(endDate,'mm/dd/yy');")
'
'    With rs
'    Do While Not .EOF
'        If Format(!startDate, "mm/dd/yy") = Format(Date, "mm/dd/yy") Then
'            printAnc = True
'        Else
'            Select Case !recur
'                Case 2
'                    printAnc = True
'
'                Case 3 'Weekly
'                    printAnc = DateDiff("d", !startDate, Date) Mod (!scope * 7) = 0
'                    'True if today(<Date>) is <scope> number of weeks from <startDate>
'
'                Case 4 'Monthly
'                    If DateDiff("m", !startDate, Date) Mod !scope = 0 Then
'                        If !dom Then
'                            printAnc = (Date = (DateAdd("d", (7 - (Weekday(DateSerial(Year(Date), Month(Date), 1)) - Weekday(!startDate))), _
'                                DateAdd("ww", !subScope - 1, DateSerial(Year(Date), Month(Date), 1)))))
'                            'IF: [1st of month + (<subscope> - 1) weeks] + [7 days - (day of week - target.day.of.week(<subscope>) ) ]
'                        Else
'                            'Fix/Verify
'                            'printAnc = (Day(Date) = Day(!subScope))
'                        End If
'                    End If
'
'                Case 5 'Yearly
'                    'Lol have fun with this one
'
'            End Select
'        End If
'
'        If printAnc Then anc = Nz(anc) & IIf(Not IsNull(!startEvent), !startEvent & " - " & IIf(Not IsNull(!endEvent), !endEvent, "UFN") & ": ", "") & !msg & vbCrLf
'        .MoveNext
'    Loop
'    End With
'    rs.Close
'    Set rs = Nothing
'
'    announcement = Nz(anc)
'    For Each ctl In Controls
'        If ctl.Tag = "announce" Then ctl.Visible = Not announcement = ""
'    Next
'
'End Function

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

Private Sub lstNotice_DblClick(cancel As Integer)
If IsNull(lstNotice) Then Exit Sub
Dim RS As DAO.Recordset
Set RS = CurrentDb.OpenRecordset("SELECT * FROM tblAnnounce WHERE ID = " & lstNotice)
With RS
    If RS.RecordCount = 1 Then
        MsgBox "Date: " & !startDate & " - " & !endDate & vbCrLf & _
                "Time: " & !startEvent & " - " & !endEvent & vbCrLf & vbCrLf & _
                !msg, _
                vbInformation, "Notice"
    End If
End With

End Sub

Private Sub new_flt_Click()
    'DoCmd.OpenForm "quick_input", , , , acFormAdd, acDialog
    'Forms!quick_input.cancel_flt.Visible = False
    
    openFlight
End Sub

Private Sub openFlight(Optional ID As Integer, Optional onStation As Boolean)
    DoCmd.OpenForm "quick_input", , , IIf(ID <> 0, "ID = " & ID, ""), IIf(ID <> 0, acFormEdit, acFormAdd), acDialog, IIf(onStation, 4, Null)
    If CurrentProject.AllForms("quick_input").IsLoaded Then
        fpReload Forms!quick_input.fpType
        DoCmd.Close acForm, "quick_input"
    End If
End Sub

Private Sub fpReload(Optional fpType As Integer)
    Select Case fpType
    Case 0
        For Each ctl In Controls
            If TypeOf ctl Is ListBox Then
                If ctl.Tag Like "*fp*" Then ctl.Requery
            End If
        Next
    
    Case 1
        inbound_aircraft.Requery
    Case 2
        outbound_aircraft.Requery
    Case 3
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
'------------------------------------------------------------
' Command131_Click
'
'------------------------------------------------------------
Private Sub Command131_Click()
On Error GoTo Command131_Click_Err

    On Error Resume Next
    DoCmd.RunCommand acCmdSelectRecord
    If (MacroError = 0) Then
        DoCmd.RunCommand acCmdCopy
    End If
    If (MacroError = 0) Then
        DoCmd.RunCommand acCmdRecordsGoToNew
    End If
    If (MacroError = 0) Then
        DoCmd.RunCommand acCmdSelectRecord
    End If
    If (MacroError = 0) Then
        DoCmd.RunCommand acCmdPaste
    End If
    If (MacroError <> 0) Then
        Beep
        MsgBox MacroError.description, vbOKOnly, ""
            End If


Command131_Click_Exit:
    Exit Sub

Command131_Click_Err:
    MsgBox Error$
    Resume Command131_Click_Exit

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

    For i = 1 To ppr_list.ListCount - 1
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

Private Sub sperr_Click()
Dim RS As DAO.Recordset
Set RS = CurrentDb.OpenRecordset("SELECT * FROM tblUserAuth WHERE username = '" & Environ$("username") & "'")
RS.edit
RS!spAccess = getAccessSP
RS.update
If RS!spAccess = True Then sperr.Visible = False
End Sub
