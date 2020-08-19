VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmDD175"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Public fpType As Integer
Dim bClose As Boolean
Dim pprSelect As Boolean

Private Sub ATActime_Click()
    DoCmd.OpenForm "check_in_out", , , "ID = " & ID, acFormEdit
End Sub

Private Sub ATDctime_Click()
    DoCmd.OpenForm "check_in_out", , , "ID = " & ID, acFormEdit
End Sub

Private Sub cancel_flt_Click()
bClose = True
    DoCmd.RunCommand acCmdSaveRecord
    If cnlFlight(Me.RecordsetClone) Then
        DoCmd.Close acForm, "quick_input", acSaveYes
        DoEvents
        DoCmd.RunCommand acCmdRefresh
    End If
End Sub

Private Sub cPPR_Change()
    pprSelect = True
End Sub

Private Sub cPPR_LostFocus()
If NewRecord Then Exit Sub
Dim RS As DAO.Recordset
Set RS = CurrentDb.OpenRecordset("SELECT * FROM tblPPR WHERE ((nz(tblPPR.PPR)<>'-') AND ((Format('" & ZToL(DOF + ETD) & "','mm/dd/yyyy')) Between Format([arrDate],'mm/dd/yyyy') And Format([depDate],'mm/dd/yyyy'))) ORDER BY tblPPR.ID;")
    
    Do While Not RS.EOF
        If Nz(PPR) = "" And depPoint <> Destination And (RS!Callsign = Callsign Or RS!Tail = Tail) And ((Destination = DLookup("station", "tblSettings") And Nz(RS!arrID, 0) = 0) Or depPoint = DLookup("station", "tblSettings")) Then
            PPR = RS!PPR
            Exit Do
        End If
        RS.MoveNext
    Loop
    
    If IsNull(DLookup("PPR", "tblPPR", "PPR = '" & PPR & "'")) Then
        MsgBox "Invalid PPR Number.", vbCritical, "Flight Plan"
        cPPR.SetFocus
    End If
End Sub

Private Sub depPoint_LostFocus()
    If IsNull(depPoint) Then depPoint = DLookup("station", "tblSettings")
End Sub

Private Sub Destination_LostFocus()
    If IsNull(Destination) Then Destination = DLookup("station", "tblSettings")
End Sub

Private Sub DOF_GotFocus()
    If Nz(DOF) = "" Then
        DOF = Format(Now, "dd-mmm-yy")
    End If
End Sub

Private Sub DOF_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
    Case 37, 40
        DOF = DateAdd("d", -1, DOF)
        KeyCode = 0
    Case 38, 39
        DOF = DateAdd("d", 1, DOF)
        KeyCode = 0
    End Select
End Sub

Private Sub done_Click()
'This needs a SERIOUS cleanup
Dim db As Database
Dim RS As DAO.Recordset
Dim rsPPR As DAO.Recordset
Dim rstSP As DAO.Recordset
Dim ances As Integer
Dim stn As String
Dim arrordep As String
Dim f As Boolean
Dim sperr As Boolean
Dim spAccess As Boolean

spAccess = DLookup("spaccess", "tbluserauth", "username = '" & getUSN & "'")
If Not spAccess Then spAccess = getAccessSP
If spAccess Then
    If Not sp Then sp = (MsgBox("Connect to SharePoint", vbQuestion + vbYesNo, "PPR") = vbYes)
    spAccess = sp
End If

stn = DLookup("Station", "tblSettings")
Set db = CurrentDb
On Error GoTo save_Click_Err

bClose = True

sending.Visible = True
sp.Enabled = False
done.Enabled = False
cancel_flt.Enabled = False
cPPR.Enabled = False
DoEvents

10  ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '''''Grammar check''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    For Each ctl In Controls
        If TypeOf ctl Is TextBox And ctl.Name <> "Remarks" Then
            ctl.Value = UCase(ctl.Value)
        End If
    Next
        
    If Nz(Tail) < 4 And Me.acType = "C172" Then
        Tail = Callsign
    End If
    
    If IsNull(Tail) Then
        MsgBox "Please enter a Tail Number.", vbCritical, "Flight Plan"
        GoTo done_fb
    Else
        If Not IsNumeric(Left(Tail, 1)) Then
'            If isMil = 0 Then
'                If MsgBox("Is this a CIVILIAN aircraft?" & vbCrLf & "(Tail number: " & Tail & ")", vbQuestion & vbYesNo, "Aircraft Type") = vbYes Then isMil = 1
'            End If
        ElseIf Not IsNumeric(Right(Tail, 1)) Then
            Tail = Left(Tail, Len(Tail) - 1)
        End If
    End If
    
    Tail = Right(Tail, 7)
    
    If Not depPoint = stn And Not Destination = stn Then
        MsgBox "If this aircraft is not landing here, a flight tracker entry is not needed." & vbCrLf & "Otherwise, please enter as two seperate flight plans.", vbInformation, "AeroStat"
        GoTo done_fb
    End If
    
    If IsNumeric(altitude) Then altitude = IIf(altitude <= 140, "A", "F") & altitude
    
    Me.acType = Replace(Me.acType, "-", "")
    Me.acType = Replace(Me.acType, "/", "")
    Me.Tail = Replace(Me.Tail, "-", "")
    
    If IsNull(DLookup("acType", "AIRCRAFT", "acType = '" & Me.acType & "'")) Then 'Type is not valid
    
        If Not IsNull(DLookup("acType", "AIRCRAFT", "synonym = '" & Me.acType & "'")) Then 'A correction can be suggested
            DoCmd.OpenForm "frmEnforceAcType", , , , , acDialog, Me.acType
            Me.acType = Forms!frmEnforceAcType.response
            DoCmd.Close acForm, "frmEnforceAcType"
            
        Else 'DB has no clue what the type is.
            MsgBox "Invalid Aircraft Type.", vbCritical, "Flight Plan"
            GoTo done_fb
        End If
    End If
    
    transient = isT(Me.Tail, Me.acType)
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
15
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '''''DOF error check''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    If Destination = stn Then
        arrDate = Format(ZToL(DOF + (Nz(ATD, ETD) + ETE)), "dd-mmm-yy") & " " & Format(ZToL(Nz(ATA, IIf(IsNull(ATD), IIf(Not ETA = 0, ETA, ETD + ETE), ATD + ETE)), True), "hh:nn")
        arrordep = arrDate
        If Not Status = "Closed" And Not DateDiff("d", arrDate, Now) = 0 Then
            If MsgBox("Calculated ETA as " & arrDate & " (L). Is this correct?" & vbCrLf & "(DOF: " & DOF & ")", vbQuestion + vbYesNo, "AeroStat") = vbNo Then
                GoTo done_fb
                Exit Sub
            End If
        End If
    End If
    If depPoint = stn Then
        depDate = ZToL(DOF + Nz(ATD, ETD))
        If arrordep = "" Then arrordep = depDate
        If Not Status = "Closed" And Not DateDiff("d", depDate, Now) = 0 Then
            If MsgBox("Calculated ETD as " & depDate & " (L). Is this correct?" & vbCrLf & "(DOF: " & DOF & ")", vbQuestion + vbYesNo, "AeroStat") = vbNo Then
                GoTo done_fb
                Exit Sub
            End If
        End If
    End If
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    If depPoint = stn Then
        If Destination = stn Then
            Me.fpType = 3
        Else
            Me.fpType = 2
        End If
    Else
        Me.fpType = 1
    End If
    If OpenArgs = 4 Then Me.fpType = 4
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '''''On Station Check'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    Set qdf = db.QueryDefs("qOnStation")
    qdf.Parameters("bos") = transient
    Set RS = qdf.OpenRecordset
    Dim edo As Boolean
    If depPoint = stn Then
        Do While Not RS.EOF
            If Right(RS!Tail, 4) = Right(Me.Tail, 4) Or RS!Callsign = Me.Callsign Then
                If RS!ID <> ID Then
                    'rs.edit
                    If IsNull(ParentID) Then ParentID = RS!ID
                    'rs!Stationed = False
                    With db
                        .Execute "UPDATE tblTraffic SET Stationed = False WHERE Stationed = True AND ID <> " & ID & " AND (Right(Tail, 4) = '" & Right(Me.Tail, 4) & "' Or Callsign = '" & Callsign & "')"
                        log "Updated " & .RecordsAffected & " records", Me.Name & ".done_Click"
                    End With
                    Stationed = True
                    PPR = RS!PPR
                    If IsNull(Spot) Then Spot = RS!Spot
                    'rs.update
                    edo = True
                    Exit Do
                End If
            End If
            RS.MoveNext
        Loop

        'If Not edo And Format(Nz(arrDate, depDate), "ddmmyy") = Format(Date, "ddmmyy") Then
        If Not edo And serialDate(Nz(arrDate, depDate)) = serialDate(Date) Then
            Set RS = db.OpenRecordset("qInbound")
            With RS
            Do While Not .EOF
                If (Right(RS!Tail, 4) = Right(Me.Tail, 4) Or !Callsign = Me.Callsign) And PPR = RS!PPR Then
                    If IsNull(Spot) Then Spot = !Spot
                    If IsNull(ParentID) Then ParentID = !ID
                    PPR = !PPR
                    Stationed = False
                    edo = True
                    Exit Do
                End If
                .MoveNext
            Loop
            End With
        End If
        
'        If Not edo Then
'            If Not Stationed Then Stationed = (MsgBox("Is this aircraft on station?" & vbCrLf & vbCrLf & "(Inbound record not found)", vbQuestion + vbYesNo, "Flight Plan") = vbYes)
'        End If
    End If


    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '''''Update PPR Tracker'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
PPR: ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    If IsNull(DLookup("PPR", "tblPPR", "PPR = '" & PPR & "'")) Or Nz(PPR) = "" Then '???
        PPR = "" 'LAG ALERT (serialDate)
        Set rsPPR = db.OpenRecordset("SELECT * FROM tblPPR WHERE tblPPR.PPR <> '-' AND Format('" & Nz(arrDate, depDate) & "','mm/dd/yyyy') Between Format(nz([arrDate],[depDate]),'mm/dd/yyyy') And Format(nz([depDate],[arrDate]),'mm/dd/yyyy') ORDER BY tblPPR.arrDate DESC;")
    Else
        Set rsPPR = db.OpenRecordset("SELECT * FROM tblPPR WHERE tblPPR.PPR = '" & PPR & "'")
    End If
    
    
    
    If Not depPoint = Destination Then
        'GoTo done_fb
        'If Destination = stn Then
        'Set rsPPR = db.OpenRecordset("SELECT * FROM tblPPR WHERE (((tblPPR.PPR)<>'-') AND ((Format(Date(),'mm/dd/yyyy')) Between Format([arrDate],'mm/dd/yyyy') And Format([depDate],'mm/dd/yyyy'))) ORDER BY tblPPR.ID;")
        
            With rsPPR: Do While Not .EOF
                If ((!Callsign = Callsign Or !Tail = Tail) Or !PPR = PPR) And !Status <> "Cancelled" Then
                    If pprSelect Or (Not pprSelect And ((Destination = stn And (!arrID = 0 Or !arrID = ID)) Xor depPoint = stn)) Then 'Found PPR based on Callsign, Tail, or previous PPR assignment
                        .edit
                        spID = rsPPR!spID
                            !arrID = ID
                        If Destination = stn Then
                            !Type = Me.acType
                            !arrID = ID
                            spID = !spID
                            PPR = !PPR
                            !Tail = Tail
                            dvCode = !dvCode
                            If Nz(Spot) = "" Then
                                Spot = !Spot
                            Else
                                !Spot = Spot
                            End If
                            
                            'db.Execute "UPDATE tblPPR SET arrID = " & ID & ", tblPPR.acType = '" & Me.acType & "', tail = '" & Tail & "' WHERE " & _
                                        "(((Callsign = '" & Callsign & "' or tail = '" & Tail & "') or ppr = '" & PPR & "') And status <> 'Cancelled')" & _
                                        "AND ((Destination = '" & stn & "' And (arrID = 0 Or arrID = " & ID & ")) Xor depPoint = '" & stn & "')"
                            

                        ElseIf depPoint = stn Or Nz(PPR) = "" Then
                            
                            PPR = !PPR
                            !Tail = Tail
                            spID = !spID
                            dvCode = !dvCode
                            If Nz(Spot) = "" Then
                                Spot = !Spot
                            Else
                                !Spot = Spot
                            End If
'                            Select Case Status: Case Is <> "Closed", Is <> "Enroute"
'                                Stationed = True
'                            End Select
                            !Type = Me.acType
                        End If
                        
                        If PPR = !PPR And !Callsign <> Callsign Then
                            If !Tail <> Tail Then
                                ans = MsgBox("PPR: " & !PPR & " / " & !Callsign & vbCrLf & vbCrLf & "Update PPR Callsign?", vbQuestion + vbYesNo, "Flight Plan")
                            Else
                                ans = vbYes
                            End If
                            If ans = vbYes Then
                                spID = !spID
                                !Type = Me.acType
                                !Remarks = !Remarks & IIf(Nz(!Remarks) = "", "", vbCrLf) & Format(Now, "dd mmm: ") & "Callsign change: " & vbCrLf & !Callsign & " -> " & Callsign
                                !Callsign = Callsign
                            End If
                        End If
                        .update
                        Exit Do
                    Else
                        .MoveNext
                    'Exit Do
                    End If
                Else
                    .MoveNext
                End If
            Loop: End With
            If transient And Nz(PPR, "-") = "-" Then
                If Not pprSelect Then
    '               If MsgBox("This flight plan is missing a PPR. Continue?", vbQuestion + vbYesNo, "Flight Plan") = vbNo Then
    '                    GoTo done_fb
    '                End If
                    DoCmd.OpenForm "frmPPRselect", , , , , acDialog, Nz(arrDate, depDate)
                    pprSelect = True
                    GoTo PPR
                Else
                    MsgBox "This flight plan requires a PPR. If this is aircraft is assigned to this base, please check the tail number, or register it to Base Assigned Aircraft.", vbInformation, "Flight Plan"
                    GoTo done_fb
                End If
            End If
        'End If
    End If
  
    If IsNumeric(Left(Me.acType, 1)) Then
        Me.Number = Left(Me.acType, 1)
        Me.acType = Right(Me.acType, Len(Me.acType) - 1)
    End If
    
    'countType = ICAOToCount(Me.acType, Me.isMil)
    countType = Me.acType
    
    If Format(ETD, "hhnn") = "0000" Then ETD = "0:01"
    
    
    
'    If IsNull(depPoint) Then depPoint = stn
'    If IsNull(Destination) Then Destination = stn
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
30  ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '''''Dupe protection''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Set RS = db.OpenRecordset("SELECT * FROM tblTraffic WHERE (tblTraffic.Callsign = '" & Callsign & "' OR tblTraffic.Tail = '" & Tail & "') AND (tblTraffic.depPoint Like '" & depPoint & "' OR tblTraffic.Destination Like '" & Destination & "') AND Format([DOF],'mmddyy') = '" & Format(DOF, "mmddyy") & "' AND (Format([ETD],'hhnn') Between '" & Format(LToZ(Nz(arrDate, depDate)), "hhnn") & "' And '" & Format(LToZ(Nz(depDate, arrDate)), "hhnn") & "' OR Format(IIf(IsNull(ATD),IIf(Not ETA=0,ETA,ETD+ETE),ATD+ETE),'hhnn') Between '" & Format(LToZ(Nz(arrDate, depDate)), "hhnn") & "' And '" & Format(LToZ(Nz(depDate, arrDate)), "hhnn") & "')")
    Do While Not RS.EOF
        If Not ID = RS!ID And RS!Status <> "Closed" Then
            If RS!Status = "Cancelled" Then
                Stationed = RS!Stationed
                RS.delete
                Exit Do
            ElseIf MsgBox("A duplicate flight plan was found. Would you like to update it?" & vbCrLf & "(This flight plan will be discarded)", vbQuestion + vbYesNo, "Update Existing?") = vbYes Then
                Undo
                DoCmd.OpenForm "quick_input", , , "ID = " & RS!ID, acFormEdit
                GoTo done_fb
            ElseIf MsgBox("This flight plan conflicts with another flight plan. Continue?", vbQuestion + vbYesNo, "Confirm") = vbNo Then
                GoTo done_fb
            End If
        End If
        RS.MoveNext
    Loop
    

    If Status = "-" Then Status = "Pending"
        

50
    

60  ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '''''Sharepoint'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    If spAccess And Not Me.acType = "C172" Then
        Dim sprecord As Date
        Dim csnospace As String
        sending.Caption = "Connecting to SharePoint...."
        DoEvents
        
        If IsNull(spID) Then
            Set qdf = CurrentDb.QueryDefs("qMissionTracker")
            qdf.Parameters("fromDay") = Format(arrordep, "mm/dd/yy")
            Set rstSP = qdf.OpenRecordset()
        Else
            Set qdf = CurrentDb.QueryDefs("qMissionTrackerByID")
            qdf.Parameters("mtbid") = spID
            Set rstSP = qdf.OpenRecordset()
            ovr = (rstSP![Mission/Event Type] = "PPR")
        End If
        
        If Not sperr Then
        'If IsNull(spID) Then
61
            sperr = True
            Do While Not rstSP.EOF
                csnospace = Replace(Nz(rstSP![Call Sign]), " ", "")
                'FIX THIS NOW!
                If (Me.Callsign = csnospace And Right(Me.Tail, 4) = Right(Nz(rstSP![Tail Number], Me.Tail), 4)) _
                Or ((Left(Callsign, 3) = Left(csnospace, 3) Or Len(csnospace) <= 3) And Me.Tail = Nz(rstSP![Tail Number])) _
                Or spID Like rstSP!ID _
                Or PPR Like Nz(rstSP![PPR #], "NOPPR") Then
                'If (Left(Me.Callsign, 3) = Left(csnospace, 3) And Me.Tail = Nz(rstSP![Tail Number])) Or ((Not Me.acType = "C30J" Or Not Me.acType = "C130")) Or ((Me.acType = "C30J" Or Me.acType = "C130") And (Right(Me.Tail, 4) = Nz(rstSP![Tail Number]))) Then
                'if csnospace = Me.Callsign Or spID = rstSP!ID Then
                    If Left(Callsign, 5) = "KANTO" Then Exit Do
                    
                    If Not IsNull(spID) Then
                        sending.Caption = "Sending to SharePoint...."
                        DoEvents
                        sperr = False
                        Exit Do
                    End If
                    If (Destination = stn And Abs(DateDiff("h", Nz(rstSP![ETA (Z)], Nz(arrDate)), LToZ(Nz(arrDate)))) <= 3) Or (depPoint = stn And Abs(DateDiff("h", Nz(rstSP![ETD (Z)], Nz(depDate)), LToZ(Nz(depDate)))) <= 3) Then
                        sending.Caption = "Sending to SharePoint...."
                        'spID = rstSP!PPR
                        DoEvents
                        sperr = False
                        Exit Do
                    End If
                End If
                rstSP.MoveNext
            Loop
62
        'Else
            If Not sperr Then
                rstSP.edit
                sending.Caption = "Sending to SharePoint...."
                DoEvents
                spID = Val(rstSP!ID)
                csnospace = Replace(Nz(rstSP![Call Sign]), " ", "")
                If csnospace <> Callsign Then rstSP![Call Sign] = Me.Callsign
                
                If PPR <> "-" Then
                With rsPPR
                    'rsPPR.Bookmark = rsPPR.LastModified
                    rstSP![PPR #] = PPR
                    If Nz(!spID, 0) = 0 Then
                        !spID = "  spID "
                        rstSP![Current ICAO] = !depPoint
                        rstSP![Next ICAO] = !pprDestination
                        rstSP![ETD (Z)] = LToZ(Nz(!depDate))
                        
                        rstSP![DV Code] = !dvCode
                        rstSP!Fuel = !Fuel
                    End If
                End With
                End If
                
                If Not rstSP![Tail Number] Like Me.Tail Then
                    
                    If MsgBox("Tail number does not match SharePoint:" & vbCrLf & "SharePoint Tail: " & Nz(rstSP![Tail Number], "None") & vbCrLf & "Your Tail: " & Tail & _
                            vbCrLf & vbCrLf & "Update SharePoint?", vbQuestion + vbYesNo, "Flight Plan") = vbYes Then
                        If PPR <> "-" Then
                            With rsPPR
                                '.Bookmark = .LastModified
                                .edit
                                rstSP![Tail Number] = Me.Tail
                                !Tail = Me.Tail
                                .update
                                .Bookmark = .LastModified
                            End With
                        End If
                    Else
                        If PPR <> "-" Then
                            With rsPPR
                                '.Bookmark = .LastModified
                                .edit
                                Me.Tail = rstSP![Tail Number]
                                !Tail = Me.Tail
                                .update
                                .Bookmark = .LastModified
                            End With
                        End If
                    End If
                End If
                
                'rstSP![Call Sign] = Callsign
63
                'Arrival
                If Destination = stn Then
                    rsPPR.Bookmark = rsPPR.LastModified
                    If Nz(Spot) = "" Then Spot = rsPPR!Spot
                    rstSP![Current ICAO] = depPoint
                    rstSP![ETA (Z)] = LToZ(arrDate)
                    If Not IsNull(ATA) Then rstSP![ATA (Z)] = LToZ(Nz(arrDate))
                    
'                    Select Case Nz(Spot)
'                        Case "", "AMC", "TBD" 'We do not have assignment
'                            If NewRecord Then
'                                Spot = IIf(Left(Nz(rstSP![Parking Spot/Location]), 3) = "HOT", "HC" & Right(Nz(rstSP![Parking Spot/Location]), 1), Nz(rstSP![Parking Spot/Location]))
'                            Else
'                                GoTo NEXTCASE
'                            End If
'                            'rstSP![Parking Spot/Location] = Spot
'
'                        Case Is <> rstSP![Parking Spot/Location] 'We have assignment, but it doesnt match the SharePoint
'NEXTCASE:
'                            If Nz(rstSP![Parking Spot/Location]) = "" Then
'
'                                Nz(rstSP![Parking Spot/Location]) = Spot
'
'                            ElseIf Nz(rstSP![Parking Spot/Location]) <> Spot Then
'                                If MsgBox("Parking assignment does not match SharePoint:" & vbCrLf & "SharePoint Spot: " & rstSP![Parking Spot/Location] & vbCrLf & "Your Spot: " & Spot & _
'                                vbCrLf & vbCrLf & "Update SharePoint?", vbQuestion + vbYesNo, "Flight Plan") = vbYes Then
'                                    rstSP![Parking Spot/Location] = Spot
'                                Else
'                                    Spot = rstSP![Parking Spot/Location]
'                                End If
'                            End If
'                    End Select
                    
                    With rstSP: Select Case Nz(Spot)
                            Case "", "AMC", "TBD"  'We do not have assignment
                                Select Case Nz(![Parking Spot/Location])
                                Case "", "AMC", "TBD"
                                    If Spot <> Nz(![Parking Spot/Location]) Then ![Parking Spot/Location] = Spot
                                Case Else
                                    GoTo els
                                End Select
                                
                                Spot = IIf(Left(Nz(![Parking Spot/Location]), 3) = "HOT", "HC" & Right(Nz(![Parking Spot/Location]), 1), Nz(![Parking Spot/Location]))
                                '![Parking Spot/Location] = Spot
                                    
                            Case Is <> ![Parking Spot/Location] 'We have assignment, but it doesnt match the SharePoint
els:
                                If Nz(![Parking Spot/Location]) = "" Then
                                
                                    ![Parking Spot/Location] = Spot
                                    
                                ElseIf MsgBox("Parking assignment does not match SharePoint:" & vbCrLf & "SharePoint Spot: " & ![Parking Spot/Location] & vbCrLf & "Your Spot: " & Spot & _
                                vbCrLf & vbCrLf & "Update SharePoint?", vbQuestion + vbYesNo, "Flight Plan") = vbYes Then
                                    ![Parking Spot/Location] = Spot
                                Else
                                    Spot = ![Parking Spot/Location]
                                End If
                    End Select: End With
                    
                    If PPR <> "-" Then
                        With rsPPR
                            '.Bookmark = .LastModified
                            .edit
                            !Spot = Spot
                            .update
                            .Bookmark = .LastModified
                        End With
                    End If
                    
                    If Not depPoint = Destination Then
                        If IsNull(rstSP![Start Time]) Then rstSP![Start Time] = Date
                        If IsNull(rstSP![End Time]) Then rstSP![End Time] = Nz(depDate, Date + 1)
                    End If
                End If
64
                'Departure
                If depPoint = stn Then
                    rstSP![Next ICAO] = Destination
                    'If LToZ(depDate) > Nz(rstSP![ETD (Z)], 0) Then rstSP![ETD (Z)] = DateAdd("h", -DLookup("Timezone", "tblSettings"), depDate)
                    If Format(LToZ(depDate), "dd") = Format(Nz(rstSP![ETD (Z)], 32), "dd") Then
                        'rstSP![ETD (Z)] = DateAdd("h", -DLookup("Timezone", "tblSettings"), depDate)
                        rstSP![ETD (Z)] = LToZ(Nz(depDate))
                        If Not IsNull(ATD) Then rstSP![ATD (Z)] = LToZ(Nz(depDate))
                    End If
                    If Not depPoint = Destination Then
                        If IsNull(rstSP![Start Time]) Then rstSP![Start Time] = Date
                        If IsNull(rstSP![End Time]) Then rstSP![End Time] = Nz(depDate, Date)
                    End If
                    If Destination <> stn And PPR <> "-" Then
                        With rsPPR
                            '.Bookmark = .LastModified
                            .edit
                            !depDate = depDate
                            .update
                            .Bookmark = .LastModified
                        End With
                    End If
                End If
                rstSP![AM Initials] = AMOPS
                If IsNull(rstSP![Start Time]) Then rstSP![Start Time] = Date
                If IsNull(rstSP![End Time]) Then rstSP![End Time] = Date
                rstSP.update
                sending.Caption = "Success!"
                DoEvents
            Else
                sending.Caption = "Not Found"
                DoEvents
            End If
            rstSP.Close
            Set rstSP = Nothing
        End If
        If sperr Then
            sending.Caption = "Not Found"
            DoEvents
        End If
        'End If
    End If
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    RS.Close
    Set RS = Nothing
    
done_exit:
    Dirty = False
    sending.Visible = False
    syncTrafficLog ID, "Traffic", NewRecord
    
    Me.Visible = False
'    If Forms!frmMain!view.SourceObject = "Traffic Log" Then
'        For Each ctl In Forms!frmMain!view.Controls
'            If TypeOf ctl Is ListBox Then ctl.Requery
'        Next
'        Forms!frmMain!view.Form.osReload
'    End If
    Exit Sub
    
done_fb:
    bClose = False
    sending.Visible = False
    sp.Enabled = True
    done.Enabled = True
    cancel_flt.Enabled = True
    cPPR.Enabled = True
    Exit Sub

save_Click_Err:
    If err = 2046 Then Resume Next
    ErrHandler err, Error$, Me.Name & ".done_Click"
    sperr = IIf(err = 3011, True, False)
    Resume Next
    GoTo done_fb
End Sub

Private Sub altitude_LostFocus()
    If IsNull(altitude) Then
        Select Case flightRule
        Case "V"
            altitude = "VFR"
        Case Else
            altitude.BorderColor = RGB(255, 0, 0)
            altitude.SetFocus
        End Select
    Else
        altitude.BorderColor = RGB(166, 166, 166)
    End If
End Sub

Private Sub Form_BeforeUpdate(cancel As Integer)
    If Not bClose Then
        'cancel = True
        Undo
        Exit Sub
    End If
End Sub

Private Sub Form_Close()
'On Error Resume Next
'    If CurrentProject.AllForms("checked_flts").IsLoaded Then
'        Forms!checked_flts.done.Requery
'        Forms!checked_flts.done.SetFocus
'    Else
'        Dim tl As Form
'        Set tl = IIf(CurrentProject.AllForms("Traffic Log").IsLoaded, Forms![traffic log], Forms!frmmain)
'
'        For Each ctl In tl
'            If TypeOf ctl Is ListBox Then ctl.Requery
'        Next
'    End If
End Sub

Private Sub Form_Current()
    bClose = False
    sp = NewRecord Or Not IsNull(spID)
    If Status = "Cancelled" Then
        For Each ctl In Controls
            If ctl.Name Like "*ctime*" Then ctl.Visible = True
            If ctl.Name Like "*clr*" Then ctl.BackColor = RGB(255, 0, 0)
            FormHeader.BackColor = RGB(255, 0, 0)
            FormFooter.BackColor = RGB(255, 0, 0)
        Next
        cancel_flt.Visible = True
        cancel_flt.Caption = "Re-Activate"
 '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ElseIf Status = "Enroute" Then
        For Each ctl In Controls
            If ctl.Name Like "*ctime*" Then ctl.Visible = True
            If ctl.Name Like "*clr*" Then ctl.BackColor = RGB(0, 255, 0)
            FormHeader.BackColor = RGB(0, 255, 0)
            FormFooter.BackColor = RGB(0, 255, 0)
            cancelled.Visible = True
            cancelled.Caption = Replace(cancelled.Caption, "Cancelled", "Enroute")
        Next
 '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ElseIf Status = "Closed" Then
        For Each ctl In Controls
            If ctl.Name Like "*ctime*" Then ctl.Visible = True
            If ctl.Name Like "*clr*" Then ctl.BackColor = RGB(190, 190, 190)
            FormHeader.BackColor = RGB(190, 190, 190)
            FormFooter.BackColor = RGB(190, 190, 190)
            cancelled.Visible = True
            cancelled.Caption = Replace(cancelled.Caption, "Cancelled", "Closed")
        Next
    Else
        cancel_flt.Visible = True
    End If
    Dim RS As DAO.Recordset
    Set RS = CurrentDb.OpenRecordset(cPPR.RowSource)
    Do While Not RS.EOF
        If PPR = RS!PPR Then
            cPPR = RS!PPR
            Exit Do
        End If
        RS.MoveNext
    Loop
    RS.Close
    Set RS = Nothing
    
    sp = (Nz(spID) <> 0 Or NewRecord)
    
    cancel_flt.Visible = Not NewRecord
    Tail.RowSource = Nz(longTail)
End Sub

Private Sub Form_Error(DataErr As Integer, response As Integer)
    If DataErr = 2169 Then response = True
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 And Shift = acShiftMask Then
        done.SetFocus
        done_Click
    End If
End Sub

Private Sub Form_Load()
'    bClose = False
'    sp = NewRecord Or Not IsNull(spID)
'    If Status = "Cancelled" Then
'        For Each ctl In Controls
'            If ctl.Name Like "*ctime*" Then ctl.Visible = True
'            If ctl.Name Like "*clr*" Then ctl.BackColor = RGB(255, 0, 0)
'            FormHeader.BackColor = RGB(255, 0, 0)
'            FormFooter.BackColor = RGB(255, 0, 0)
'        Next
'        cancel_flt.Visible = True
'        cancel_flt.Caption = "Re-Activate"
' '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'    ElseIf Status = "Enroute" Then
'        For Each ctl In Controls
'            If ctl.Name Like "*ctime*" Then ctl.Visible = True
'            If ctl.Name Like "*clr*" Then ctl.BackColor = RGB(0, 255, 0)
'            FormHeader.BackColor = RGB(0, 255, 0)
'            FormFooter.BackColor = RGB(0, 255, 0)
'            cancelled.Visible = True
'            cancelled.Caption = Replace(cancelled.Caption, "Cancelled", "Enroute")
'        Next
' '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'    ElseIf Status = "Closed" Then
'        For Each ctl In Controls
'            If ctl.Name Like "*ctime*" Then ctl.Visible = True
'            If ctl.Name Like "*clr*" Then ctl.BackColor = RGB(190, 190, 190)
'            FormHeader.BackColor = RGB(190, 190, 190)
'            FormFooter.BackColor = RGB(190, 190, 190)
'            cancelled.Visible = True
'            cancelled.Caption = Replace(cancelled.Caption, "Cancelled", "Closed")
'        Next
'    Else
'        cancel_flt.Visible = True
'    End If
'    Dim rs As DAO.Recordset
'    Set rs = CurrentDb.OpenRecordset(cPPR.RowSource)
'    Do While Not rs.EOF
'        If PPR = rs!PPR Then
'            cPPR = rs!PPR
'            Exit Do
'        End If
'        rs.MoveNext
'    Loop
'    rs.close
'    Set rs = Nothing
'
'    cancel_flt.Visible = Not NewRecord
'    Tail.RowSource = Nz(longTail)
End Sub

Private Sub openPPR_Click()
    If Nz(PPR) = "" Then Exit Sub
    DoCmd.OpenForm "new_ppr", , , "PPR = '" & PPR & "'", acFormEdit
End Sub

Private Sub roundRobin_AfterUpdate()
    If depPoint <> Destination And roundRobin Then
        roundRobin = False
    End If
    If roundRobin Then stereoPlan.Dropdown
End Sub

Private Sub Tail_GotFocus()
On Error GoTo errtrap
Dim res() As String

If Number > 1 And NewRecord Then
    res = Split(InputBox("Please enter the tail numbers, separated by spaces:" & vbCrLf & "(Ex: 1234 5678 1212...)", "Enter Tails"), " ")
    For Each t In res: Tail.AddItem t: Next
    Tail = Nz(res(0))
    longTail = Tail.RowSource
End If

sexit:
    Exit Sub
errtrap:
    Debug.Print "(" & err & "): " & Error$
End Sub

