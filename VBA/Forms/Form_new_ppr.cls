VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_new_ppr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Public bClose, fighter, frmRestrictionsResponse, frmRestrictionsResponseDoNotShow, bDELETE, dupeCheckFail, saveResult As Boolean
Dim pprGenerator, rsDUPE As String
Dim unsavedComment As Integer
Dim gFail As Boolean

Private Sub approach_AfterUpdate()
    If Not approach And Spot Like "APCH" Then Spot = ""
End Sub

Private Sub approval_Click()
    Select Case approval
        Case 1 'Approved
            FormHeader.BackColor = RGB(0, 255, 0)
            FormFooter.BackColor = RGB(0, 255, 0)
            Me.Status = "Approved"
        Case -1 'Denied
            FormHeader.BackColor = Toggle56.BackColor
            FormFooter.BackColor = Toggle56.BackColor
            Me.Status = "Denied"
        Case 2 'Requested
            FormHeader.BackColor = Toggle57.BackColor
            FormFooter.BackColor = Toggle57.BackColor
            Me.Status = "Requested/Pending"
        Case 0 'Cancelled
            FormHeader.BackColor = Toggle58.BackColor
            FormFooter.BackColor = Toggle58.BackColor
            Me.Status = "Cancelled"
    End Select
End Sub

Private Sub arrDate_AfterUpdate()
    
    frmRestrictionsResponse = False
    If Not frmRestrictionsResponseDoNotShow Then 'If the check box in frmRestrictions was checked
    'I'm sure theres a better way to write this...
        rcheck = False
        Dim cd1, cd2 As Date
        Set qdf = CurrentDb.QueryDefs("qAnnounce")
        cd1 = Format(arrDate, "mm/dd/yyyy")
        cd2 = Format(Nz(depDate, arrDate), "mm/dd/yyyy")
        qdf.Parameters("checkDate") = cd1
        Set rsRes = qdf.OpenRecordset
        With rsRes
            If Not .EOF Then
    '            .MoveLast
    '            .MoveFirst
                rcheck = True
                DoCmd.OpenForm "frmRestrictions", , , , , acDialog, cd1
                If frmRestrictionsResponse Then
                    arrDate.SetFocus
                End If
            End If
        End With
        
        
    '    If cd1 <> cd2 Then
    '        qdf.Parameters("checkDate") = cd2
    '        Set rsRes = qdf.OpenRecordset
    '        With rsRes
    '            If Not .EOF Then
    '    '            .MoveLast
    '    '            .MoveFirst
    '                DoCmd.OpenForm "frmRestrictions", , , , , acDialog, cd2
    '                If resOverride Then GoTo fallback
    '            End If
    '        End With
    '    End If
    End If

sexit:
    formatPPR
End Sub

Private Function formatPPR() As String
'PPR Generator
Dim gen As String
gen = Format(Nz(arrDate, Now), "yy") & "-" & AMOPS & _
            Format(Nz(DMax("right([ppr],3)", "tblppr", _
            "left([ppr],2) = '" & Format(Nz(arrDate, Now), "yy") & "' AND mid([ppr],instr(1,[ppr],'-') + 1,2) = '" & AMOPS & "'"), 0) + 1, "000")
    
    If pprGenerator = Nz(formPPR) Then
        pprGenerator = gen
        formPPR = gen
    End If
    
    formatPPR = gen
End Function

Private Sub btnClone_Click()
If NewRecord Then Exit Sub
Dim rsi As DAO.Recordset
Dim rss As DAO.Recordset
'On Error Resume Next

    For Each ctl In Controls
        If TypeOf ctl Is TextBox And ctl.Name <> "internalComments" And ctl.Name <> "Remarks" Then ctl.Value = UCase(ctl.Value)
    Next
    
    If (approval <= 0 Xor (Status = "Cancelled" Or Status = "Denied")) Then 'First time I've ever used a XOR gate in my life. #TooSmart
        If MsgBox(IIf(Not (Status = "Cancelled" Or Status = "Denied"), "Cancel PPR?", "Re-activate?"), vbQuestion + vbYesNo, "AeroStat") = vbYes Then
            Status = IIf(Not Status = "Cancelled", "Cancelled", "Approved")
            If Not savePPR() Then
                Exit Sub
            End If
        End If
    ElseIf Not savePPR Then
        Exit Sub
    End If
    
    Select Case approval
        Case 1
            Status = "Approved"
        Case -1
            Status = "Denied"
        Case 2
            Status = "Pending"
        Case 0
            Status = "Cancelled"
    End Select
    
    'TODO
    
    
    DoCmd.RunCommand acCmdCopy
    Dirty = False
    syncTrafficLog ID, "tblPPR", NewRecord
    
    
    DoCmd.Close acForm, "new_ppr"
    DoCmd.OpenForm "new_ppr", , , , acFormAdd, acDialog
    DoCmd.RunCommand acCmdPaste
    
    
End Sub

Private Sub btnDeleteComment_Click()
If Nz(internalComments, 0) = 0 Then Exit Sub
Dim RS As DAO.Recordset
Set RS = CurrentDb.OpenRecordset("SELECT * FROM tblPPRComments WHERE ID = " & internalComments)
With RS
    .edit
    !archive = True
    !archiveTime = Now
    !archiveBy = Util.getOpInitials(getUSN)
    .update
    .Close
End With
Set RS = Nothing
internalComments.Requery
End Sub

Private Sub btnNewComment_Click()
Dim RS As DAO.Recordset
Set RS = CurrentDb.OpenRecordset("tblPPRComments")
Dim comment As String
comment = InputBox("Please enter a comment:", "PPR", "Updated information")
If comment = "" Then Exit Sub

    With RS
        .AddNew
        !PPR = formPPR
        !inputTime = Now
        !comment = comment
        !opInitials = Util.getOpInitials(getUSN)
        .update
        .Bookmark = .LastModified
        If NewRecord Then unsavedComment = !ID
        .Close
    End With
    Set RS = Nothing
    internalComments.Requery
    
End Sub

Private Sub Command69_Click()
    DoCmd.OpenForm "frmAMC"
End Sub

Private Sub depDate_AfterUpdate()
    
    If IsDate(depDate) Then
        frmRestrictionsResponse = False
        If Not frmRestrictionsResponseDoNotShow Then 'If the check box in frmRestrictions was checked
            'I'm sure theres a better way to write this...
            rcheck = False
            Dim cd1, cd2 As Date
            Set qdf = CurrentDb.QueryDefs("qAnnounce")
            cd1 = Format(arrDate, "mm/dd/yyyy")
            cd2 = Format(Nz(depDate, arrDate), "mm/dd/yyyy")
            qdf.Parameters("checkDate") = cd1
            Set rsRes = qdf.OpenRecordset
        '    With rsRes
        '        If Not .EOF Then
        ''            .MoveLast
        ''            .MoveFirst
        '            rcheck = True
        '            DoCmd.OpenForm "frmRestrictions", , , , , acDialog, cd1
        '            If frmRestrictionsResponse Then
        '                arrDate.SetFocus
        '                arrDate = ""
        '            End If
        '        End If
        '    End With
            
            If cd1 <> cd2 Then
                qdf.Parameters("checkDate") = cd2
                Set rsRes = qdf.OpenRecordset
                With rsRes
                    If Not .EOF Then
            '            .MoveLast
            '            .MoveFirst
                        DoCmd.OpenForm "frmRestrictions", , , , , acDialog, cd2
                        If frmRestrictionsResponse Then
                            depDate.SetFocus
                        End If
                    End If
                End With
            End If
        End If
    End If

sexit:

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 And Shift = acShiftMask Then
        save.SetFocus
        save_Click
    End If
End Sub

Private Sub Form_Load()
'    Select Case approval
'        Case 1 'Approved
'            FormHeader.BackColor = RGB(0, 255, 0)
'            FormFooter.BackColor = RGB(0, 255, 0)
'        Case -1 'Denied
'            FormHeader.BackColor = Toggle56.BackColor
'            FormFooter.BackColor = Toggle56.BackColor
'        Case 2 'Pending
'            FormHeader.BackColor = Toggle57.BackColor
'            FormFooter.BackColor = Toggle57.BackColor
'        Case 0 'Cancelled
'            FormHeader.BackColor = Toggle58.BackColor
'            FormFooter.BackColor = Toggle58.BackColor
'    End Select
'    approach = (Nz(Spot) = "APCH")
'
'    If NewRecord Then
'    '''''PPR Generator'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'        PPR = Format(Nz(arrDate, Now), "yy") & "-" & AMOPS & _
'            Format(Nz(DMax("right([ppr],3)", "tblppr", _
'            "left([ppr],2) = '" & Format(Nz(arrDate, Now), "yy") & "' AND mid([ppr],instr(1,[ppr],'-') + 1,2) = '" & AMOPS & "'"), 0) + 1, "000")
'        pprGenerator = PPR
'    End If
'    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'NAMT and above can change the PPR number
    If DLookup("authlevel", "tbluserauth", "username='" & Environ$("username") & "'") < 5 Or DLookup("opinitials", "tbluserauth", "username='" & Environ$("username") & "'") Like "XD" Then 'And me, LOL
        formPPR.enabled = True
        formPPR.Locked = False
    End If
    
    On Error Resume Next
    If OpenArgs = "frmPPRSelect" Then
        With Forms!quick_input
            For Each ctl In .Controls
                If ctl.Name = "Callsign" Or _
                ctl.Name = "Type" Or _
                ctl.Name = "Tail" Or _
                ctl.Name = "depPoint" Or _
                ctl.Name = "arrDate" Or _
                ctl.Name = "Spot" Then
                    Me.Controls(ctl.Name) = ctl.Value
                End If
            Next
        End With
    End If
    
    'If NewRecord Then MsgBox "Do not issue PPRs without a valid callsign or a tail number.", vbInformation, "Do not issue PPRs without a valid callsign or a tail number."
    
End Sub

Private Sub internalComments_DblClick(cancel As Integer)
If IsNull(internalComments) Then Exit Sub
MsgBox "Time: " & internalComments.Column(1) & vbCrLf & _
        "Comment: " & internalComments.Column(2) & vbCrLf & vbCrLf & _
        "By: " & internalComments.Column(3), vbInformation, "PPR Comment"
End Sub

Private Sub lstRelatedFlights_DblClick(cancel As Integer)
If Nz(lstRelatedFlights) = "" Then Exit Sub
DoCmd.OpenForm "quick_input", , , "ID = " & lstRelatedFlights, acFormEdit

End Sub

Private Sub parkingMap_Click()
DoCmd.OpenForm "frmParkingMap", , , , , , "PPR"
End Sub

Private Sub save_Click()
On Error GoTo save_err

    For Each ctl In Controls
        If TypeOf ctl Is TextBox And ctl.Name <> "internalComments" And ctl.Name <> "Remarks" Then ctl.Value = UCase(ctl.Value)
    Next
    
    Select Case approval
        Case 1
            Status = "Approved"
        Case -1
            Status = "Denied"
        Case 2
            Status = "Pending"
        Case 0
            Status = "Cancelled"
    End Select
    
    If (approval <= 0 Xor (Status = "Cancelled" Or Status = "Denied")) Then 'First time I've ever used a XOR gate in my life. #TooSmart
        If MsgBox(IIf(Not (Status = "Cancelled" Or Status = "Denied"), "Cancel PPR?", "Re-activate?"), vbQuestion + vbYesNo, "AeroStat") = vbYes Then
            Status = IIf(Not Status = "Cancelled", "Cancelled", "Approved")
            saveResult = savePPR() 'cancel
            If Not saveResult Then
                Exit Sub
            End If
        End If
    Else
        saveResult = savePPR
        If Not saveResult Then Exit Sub
    End If
    
    If Dirty Then Dirty = False
    syncTrafficLog ID, "tblPPR", NewRecord
    Me.Visible = False
    
'    If MsgBox("Email request?", vbQuestion + vbYesNo, "PPR") = vbYes Then
'        Dim subj, body As String
'
'        subj = DLookup("pprEmailSubject", "tblsettings")
'        subj = Replace(subj, "[ppr]", PPR)
'        subj = Replace(subj, "[callsign]", UCase(Callsign))
'        subj = Replace(subj, "[type]", UCase(Me.Type))
'        subj = Replace(subj, "[arrdate]", UCase(Format(arrDate, "d mmm yy")))
'
'        body = DLookup("pprEmailBody", "tblsettings")
'        body = Replace(body, "[ppr]", PPR)
'        body = Replace(body, "[callsign]", UCase(Callsign))
'        body = Replace(body, "[type]", UCase(Me.Type))
'        body = Replace(body, "[arrdate]", UCase(Format(arrDate, "d mmm yy")))
'
'        DoCmd.OpenReport "rPPR", acViewPreview, , "arrDate = #" & arrDate & "#", acHidden, Me.ID
'        DoCmd.SendObject acSendReport, "rPPR", acFormatPDF, , , , subj, body, True
'        DoCmd.Close acReport, "rPPR"
'    End If
    
    'DoCmd.RunCommand acCmdRefresh
        
save_exit:
'    If Forms!frmMain!view.SourceObject = "Traffic Log" Then
'        For Each ctl In Forms!frmMain!view.Controls
'            If TypeOf ctl Is ListBox Then ctl.Requery
'        Next
'        Forms!frmMain!view.Form.osReload
'    End If
    Exit Sub
    
save_err:
    Resume Next
End Sub


'------------------------------------------------------------
' cancel_Click
'
'------------------------------------------------------------
'Private Sub cancel_Click()
'On Error Resume Next
'bClose = True
'Dim rs As DAO.Recordset
'Dim qdf As DAO.QueryDef
'
'    If MsgBox(IIf(Not Status = "Cancelled", "Cancel PPR?", "Re-activate?"), vbQuestion + vbYesNo, "AeroStat") = vbYes Then
'        Status = IIf(Not Status = "Cancelled", "Cancelled", "Approved")
'        If savePPR(True) Then
'            DoCmd.close acForm, "new_ppr"
'            DoCmd.RunCommand acCmdRefresh
'        End If
'    End If
'
'
'cancel_Click_Exit:
'    Exit Sub
'
'cancel_Click_Err:
'    MsgBox Error$
'    Resume cancel_Click_Exit
'
'End Sub

Private Sub Form_BeforeUpdate(cancel As Integer)
    If Not bClose Then
        cancel = True
        Undo
        If NewRecord And unsavedComment <> 0 Then
            internalComments = unsavedComment
            CurrentDb.Execute "DELETE * FROM tblPPRComments WHERE ID = " & unsavedComment
        End If
        Exit Sub
    End If
    
End Sub

Private Sub Form_Current()
bClose = False
'unArchive.Visible = archive
    Select Case Status
        Case "Approved" 'Approved (1)
            FormHeader.BackColor = RGB(0, 255, 0)
            FormFooter.BackColor = RGB(0, 255, 0)
        Case "Denied" 'Denied (-1)"
            FormHeader.BackColor = Toggle56.BackColor
            FormFooter.BackColor = Toggle56.BackColor
        Case "Pending" 'Pending (2)
            FormHeader.BackColor = Toggle57.BackColor
            FormFooter.BackColor = Toggle57.BackColor
        Case "Cancelled" 'Cancelled(0)
            FormHeader.BackColor = Toggle58.BackColor
            FormFooter.BackColor = Toggle58.BackColor
    End Select
    approach = (Nz(Spot) = "APCH")
    sp = NewRecord Or Not IsNull(spID)
    
    internalComments.Requery

    If NewRecord Then formatPPR
End Sub

Private Sub Form_Error(DataErr As Integer, response As Integer)
    Select Case DataErr
    Case 2169
        response = True
    Case 2501
        DoCmd.Close acForm, Me.Name
    End Select
    
End Sub

Private Function qualityCheck() As Boolean
Dim fail As Boolean
qualityCheck = True

With ctl

    For Each ctl In Me.Controls
        fail = False
        Select Case .Name
            
            Case "Callsign", "Tail"
                fail = Nz(Callsign) = "" And Nz(Me.Tail) = ""
                
            Case "Type", "depPoint", "arrDate", "pocName", "ctcInfo", "PPR", "Spot"
                fail = Nz(.Value) = ""
                
            Case "Fuel"
                If Not approach Then fail = Nz(.Value) = ""
                
'            Case "crew", "paxOn", "paxOff", "cargoOn", "cargoOff"
'                If Not approach Then fail = Nz(.Value) = ""
                
            Case "dvCode"
                If Nz(.Value) <> "" Then fail = (Not Left(.Value, 1) = "A" And Not Left(.Value, 1) = "R" And _
                                                Not Left(.Value, 1) = "M" And Not Left(.Value, 1) = "V" And _
                                                Not Left(.Value, 1) = "S" And Not Left(.Value, 1) = "F") _
                                                Or (Not IsNumeric(Right(.Value, 1)) Or Right(.Value, 1) > 7)
        End Select
        
        If fail Then
            .BorderColor = vbRed
            qualityCheck = False
        Else
            .BorderColor = RGB(166, 166, 166)
        End If
    
    Next ctl

End With
End Function

Public Function savePPR(Optional ByVal spEnabled As Boolean) As Boolean
On Error GoTo errtrap
spEnabled = DLookup("data", "tblSettings", "key = ""sharepoint""")
If spEnabled Then spEnabled = btnSP
    
    If Not qualityCheck Then Exit Function
    
    sending.Visible = True
    sending.Caption = "Saving..."
    DoEvents
    
    'Begin SharePoint
    Do While spEnabled
        Dim db As DAO.Database: Set db = CurrentDb
        Dim RS As DAO.Recordset: Set RS = db.OpenRecordset("spPPRLog")
        With RS
        
            If NewRecord Then
                .AddNew
            Else
            
                .FindFirst "[ID] = " & spID
                If .EOF Then
                    log "Sharepoint ID did not point to record (EOF) " & spID, Me.Name & ".savePPR", "WARN"
                    Exit Do
                End If
                
                .edit
            End If
            
            sending.Caption = "Sending to SharePoint..."
            DoEvents
            
            Dim fld: For Each fld In .Fields
                Dim e As Variant: e = DLookup("solution", "tblSPConversion", "SPField = '[" & fld.Name & "]'")
                If Not IsNull(e) Then fld.Value = Eval(e)
            Next fld
            
            .update
            .Bookmark = .LastModified
            If NewRecord Then spID = !ID
            sending.Caption = "Success!"
            DoEvents
            log "PPR sent to SharePoint! (" & spID & ")", Me.Name & ".savePPR"
            RS.Close
            Set RS = Nothing
            Set db = Nothing
        Exit Do
        End With
    Loop
    
    'End SharePoint
    
    bClose = True
    If Dirty Then Dirty = False
    
    If NewRecord Then MsgBox "Callsign: " & Callsign & vbCrLf & _
            "Type: " & Me.Type & vbCrLf & _
            "Arrival Date: " & Format(arrDate, "dd-mmm-yy @ hh:nn") & vbCrLf & _
            "Departure Date: " & Format(Nz(depDate), "dd-mmm-yy @ hh:nn") & vbCrLf & _
            "Spot: " & IIf(Spot = "TBD", "NOT ASSIGNED!", Spot) & vbCrLf & _
            vbCrLf & _
            "The PPR number for this mission is: " & formPPR _
            , vbInformation, "New PPR"
    savePPR = True
        
fexit:
    sending.Visible = False
    Exit Function
errtrap:
    ErrHandler err, Error$, "new_ppr.savePPR"
End Function

'Private Function savePPRold(Optional cancelled As Boolean) As Boolean
'On Error GoTo save_err
'Dim sperr As Boolean
'Dim spAccess As Boolean
'Dim stn As String
'stn = DLookup("data","tblSettings","key = 'station'")
'spAccess = DLookup("spaccess", "tbluserauth", "username = '" & getUSN & "'")
'If Not spAccess Then spAccess = getAccessSP
'If spAccess Then
'    If Not sp Then sp = (MsgBox("Connect to SharePoint", vbQuestion + vbYesNo, "PPR") = vbYes)
'    spAccess = sp
'End If
''If spAccess Then spAccess = Not approach And Not IsNull(spID)
'
'Dim RS As DAO.Recordset
'Dim rsFP As DAO.Recordset
'Dim rsRes As DAO.Recordset
'Dim rrstr As String
'Dim qdf As DAO.QueryDef
'save.enabled = False
'approval.enabled = False
'DoEvents
'10
'
'    If Fuel Like "N/A" Or Fuel Like "NA" Then Fuel = "0"
'    If Spot Like "N/A" Or Spot Like "NA" Then Spot = ""
'    If Tail Like "N/A" Then Tail = ""
'
'    If depPoint = stn Or Destination = stn Then
'        MsgBox "Previous or Next Station cannot be this station.", vbInformation, "PPR"
'        If depPoint = stn Then depPoint = ""
'        If Destination = stn Then Destination = ""
'        GoTo fallback
'    End If
'
'    If Nz(arrDate) > Nz(depDate) And Not IsNull(depDate) Then
'        MsgBox "Departure Date cannot be before the Arrival Date.", vbCritical, "PPR"
'        GoTo fallback
'    End If
'
'    If ((IsNumeric(Left(Tail, 1)) And Len(Tail) > 5) Or (Not IsNumeric(Left(Tail, 1)) And Len(Tail) > 6)) And Not IsNumeric(Right(Tail, 1)) Then Tail = Left(Tail, Len(Tail) - 1)
'
'    Dim fail(3) As Boolean
'    For Each ctl In Controls
'        If TypeOf ctl Is TextBox Or TypeOf ctl Is ComboBox Then
'            Select Case ctl.Name
'
'                Case "Callsign"
'                    If (Not isCallsign Or Nz(ctl.Value) = "") And fail(2) Then
'                        ctl.BorderColor = RGB(255, 0, 0)
'                        fail(1) = True
'                    Else
'                        ctl.BorderColor = RGB(166, 166, 166)
'                    End If
'
'                Case "Tail"
'                    If Nz(ctl.Value) = "TBD" Then ctl.Value = ""
'                    If Nz(ctl.Value) = "" And fail(1) Then
'                        ctl.BorderColor = RGB(255, 0, 0)
'                        fail(2) = True
'                    Else
'                        ctl.BorderColor = RGB(166, 166, 166)
'                    End If
'
'                Case "arrDate", "Remarks"
'                    If Nz(ctl.Value) = "" Then
'                        ctl.BorderColor = RGB(255, 0, 0)
'                        fail(3) = True
'                    Else
'                        ctl.BorderColor = RGB(166, 166, 166)
'                    End If
'
'                Case "depDate"
'                    If Nz(ctl.Value) = "" Then ctl.Value = DateAdd("d", 3, arrDate)
'
'                Case "Fuel", "Spot"
'                    Do 'Start loop
'                        If InStr(1, ctl.Value, "-") <> 0 Then 'Finds a hyphen
'
'                            If Not (IsNumeric(Mid(ctl.Value, InStr(1, ctl.Value, "-") - 1, 1)) _
'                            And IsNumeric(Mid(ctl.Value, InStr(1, ctl.Value, "-") + 1, 1))) Then 'both sides of hyphen are not both numeric
'
'                                ctl.Value = Replace(ctl.Value, "-", "", , 1) 'Remove hyphen; run loop again
'
'                            Else 'Both sides of hyphen are both numeric; stop the loop
'
'                                'Should probably add a check for hyphens appearing after a 'non-qualifying' hyphen, but I don't know if this situation will ever happen
'                                Exit Do
'                            End If
'                        Else 'No hyphen to begin with; prevent infinite loop
'                            Exit Do
'                        End If
'                    Loop
'
'                    If Nz(ctl.Value) = "" And Not approach Then
'                        ctl.BorderColor = RGB(255, 0, 0)
'                        fail(3) = True
'                    Else
'                        ctl.BorderColor = RGB(166, 166, 166)
'                    End If
'            End Select
'        End If
'    Next
'    If fail(1) And fail(2) Then
'        MsgBox "A Callsign or Tail Number is required to issue this PPR.", vbCritical, "PPR"
'        GoTo fallback
'    ElseIf fail(3) Then
'        GoTo fallback
'    End If
'
'    'Why is this here?
'    If NewRecord Then formatPPR
'
'    If approach Then
'        Spot = "APCH"
'        Fuel = ""
'    End If
'
''    If (Not dvCode = "" And Not IsNumeric(Right(dvCode, 1))) Or _
''    Not Left(dvCode, 1) = "A" And Not Left(dvCode, 1) = "R" And Not Left(dvCode, 1) = "M" And Not Left(dvCode, 1) = "V" And Not Left(dvCode, 1) = "S" And Not Left(dvCode, 1) = "F" Then
''        MsgBox "Invalid DV Code", vbInformation, "AeroStat"
''        Exit Function
''    End If
'
'    If Nz(dvCode) <> "" Then
'        If IsNumeric(Right(dvCode, 1)) And Right(dvCode, 1) <= 7 Then
'
'            Select Case Left(dvCode, 1)
'            Case "B", "D" To "E", "G" To "L", "N" To "Q", "T", "U", "W" To "Z"
'                MsgBox "Invalid DV Code", vbInformation, "AeroStat"
'                GoTo fallback
'            End Select
'
'        Else
'            MsgBox "Invalid DV Code", vbInformation, "AeroStat"
'            GoTo fallback
'        End If
'    End If
'
'
'    If Format(arrDate, "hhnn") = "0000" Or Format(Nz(depDate), "hhnn") = "0000" Then
'        If MsgBox("One of the times was entered as '0:00'. Is this correct?", vbQuestion + vbYesNo, "AeroStat") = vbNo Then GoTo fallback
'    End If
'
'    sending.Visible = True
'    DoEvents
'
'    For Each ctl In Controls
'        If TypeOf ctl Is TextBox Then
'            If ctl.Name = "Type" Then
'                ctl = Replace(ctl, "-", "")
'            End If
'        End If
'    Next
'20
''    rrstr = "SELECT ID, iif(startDate = endDate,startDate,startDate & ' - ' & endDate) AS [Date Range], iif((startEvent & ' - ' & endEvent) = ' - ','All Day',startEvent & ' - ' & endEvent) AS [Time Range], msg FROM tblAnnounce WHERE (Format('" & arrDate & "','mm/dd/yy') Between Format(startDate,'mm/dd/yy') And Format(endDate,'mm/dd/yy')) " & _
'                            "AND (" & _
'                                "(startDate = Format('" & arrDate & "','mm/dd/yy')) " & _
'                                "OR (recur = 2) " & _
'                                "OR (recur = 3 And (DateDiff('d',startDate,Format('" & arrDate & "','mm/dd/yy')) Mod (scope*7)=0)) " & _
'                                "OR (recur = 4 And (DateDiff('m',startDate, Format('" & arrDate & "','mm/dd/yy')) Mod scope = 0) And (" & _
'                                    "(dom = True And (Day(subscope) = Day(Format('" & arrDate & "','mm/dd/yy')))) " & _
'                                    "Or (dom = False And Format('" & arrDate & "','mm/dd/yy') = Format((DateAdd('d',base7(Weekday(startDate) - Weekday(DateSerial(Year(Format('" & arrDate & "','mm/dd/yy')), Month(Format('" & arrDate & "','mm/dd/yy')), 1))),DateAdd('ww', subScope - 1, DateSerial(Year(Format('" & arrDate & "','mm/dd/yy')), Month(Format('" & arrDate & "','mm/dd/yy')), 1)))),'mm/dd/yy')))" & _
'                                ")" & _
'                            ") OR (Format('" & Nz(depDate, 0) & "','mm/dd/yy') Between Format(startDate,'mm/dd/yy') And Format(endDate,'mm/dd/yy')) " & _
'                            "AND (" & _
'                                "(startDate = Format('" & Nz(depDate, 0) & "','mm/dd/yy')) " & _
'                                "OR (recur = 2) " & _
'                                "OR (recur = 3 And (DateDiff('d',startDate,Format('" & Nz(depDate, 0) & "','mm/dd/yy')) Mod (scope*7)=0)) " & _
'                                "OR (recur = 4 And (DateDiff('m',startDate, Format('" & Nz(depDate, 0) & "','mm/dd/yy')) Mod scope = 0) And (" & _
'                                    "(dom = True And (Day(subscope) = Day(Format('" & Nz(depDate, 0) & "','mm/dd/yy')))) " & _
'                                    "Or (dom = False And Format('" & Nz(depDate, 0) & "','mm/dd/yy') = Format((DateAdd('d',base7(Weekday(startDate) - Weekday(DateSerial(Year(Format('" & Nz(depDate, 0) & "','mm/dd/yy')), Month(Format('" & Nz(depDate, 0) & "','mm/dd/yy')), 1))),DateAdd('ww', subScope - 1, DateSerial(Year(Format('" & Nz(depDate, 0) & "','mm/dd/yy')), Month(Format('" & Nz(depDate, 0) & "','mm/dd/yy')), 1)))),'mm/dd/yy')))" & _
'                            "))"
''MOVED DUPES TO DIFFERENT SUB
'
''TODO: Show duplicates to user
'    '''''Dupe protection''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'    'Set rs = CurrentDb.OpenRecordset("SELECT * FROM tblPPR WHERE tblPPR.Callsign = '" & Callsign & "' AND ((Format(tblPPR.arrDate,'mm/dd/yy hh:nn') Between Format('" & arrDate & "', 'mm/dd/yy hh:nn') And Format('" & nz(depdate) & "', 'mm/dd/yy hh:nn')) Or (Format(tblPPR.nz(depdate),'mm/dd/yy hh:nn') Between Format('" & arrDate & "', 'mm/dd/yy hh:nn') And Format('" & nz(depdate) & "', 'mm/dd/yy hh:nn')))")
'    'Set rsDUPE = CurrentDb.OpenRecordset("SELECT * FROM tblPPR WHERE Status <> 'Cancelled' AND Callsign like '" & Callsign & "' OR (Tail not Like '' AND Tail Like '" & Tail & "');")
''    If rsDUPE.RecordCount >= 1 And Status <> "Cancelled" Then
''        With rs: Do While Not .EOF
''            If !Status <> "Cancelled" Then
'''                If (Format(!arrDate, "dd-mmm-yy hh:nn") > Format(arrDate, "dd-mmm-yy hh:nn") And Format(!arrDate, "dd-mmm-yy hh:nn") < Format(Nz(depDate), "dd-mmm-yy hh:nn")) _
'''                Or (Format(!Nz(depDate), "dd-mmm-yy hh:nn") > Format(arrDate, "dd-mmm-yy hh:nn") And Format(!Nz(depDate), "dd-mmm-yy hh:nn") < Format(Nz(depDate), "dd-mmm-yy hh:nn")) Then
''
''                If (!arrDate >= arrDate And !arrDate <= Nz(depDate)) Or (Nz(!depDate) >= arrDate And Nz(!depDate) <= Nz(depDate)) Then
''
''                    If Not ID = !ID Then
''                        If MsgBox("An overlapping PPR request was found. Would you like to view it?" & vbCrLf & "(This request will be discarded)", vbQuestion + vbYesNo, "Duplicate Found") = vbYes Then
''                            Undo
''                            DoCmd.OpenForm "new_ppr", , , "ID = " & rs!ID, acFormEdit
''                            GoTo fallback
''
''                        ElseIf MsgBox("You are about to save a PPR that overlaps with an identical request. Continue?", vbQuestion + vbYesNo, "Duplicate Found") = vbNo Then
''                            GoTo fallback
''
''                        End If
''
''                        Exit Do
''                    End If
''                End If
''            End If
''            .MoveNext
''        Loop: End With
'
''        If Not ID = rs!ID Then
''            If MsgBox("An overlapping PPR request was found. Would you like to view it?" & vbCrLf & "(This request will be discarded)", vbQuestion + vbYesNo, "Duplicate Found") = vbYes Then
''                Undo
''                DoCmd.OpenForm "new_ppr", , , "ID = " & rs!ID, acFormEdit
''                GoTo fallback
''            ElseIf MsgBox("You are about to save a PPR that overlaps with an identical request. Continue?", vbQuestion + vbYesNo, "Duplicate Found") = vbNo Then
''                GoTo fallback
''            End If
''        End If
'
'        Dim rsDUPE As DAO.Recordset 'Complicated Dupe Finder Code.....
''        Set rsDUPE = CurrentDb.OpenRecordset("SELECT * " & _
''        "FROM tblPPR " & _
''        "WHERE tblPPR.Status <> 'Cancelled' and tblPPR.archive = False and tblPPR.PPR <> '" & PPR & "'" & _
''        "AND (tblPPR.Callsign = '" & Callsign & "' Or (nz(tblPPR.Tail) <> '' And tblPPR.Tail = '" & Nz(Tail) & "')   ) " & _
''        "AND ((Format(tblPPR.arrDate,'mm/dd/yy hh:nn') Between Format('" & arrDate & "', 'mm/dd/yy hh:nn') And Format('" & Nz(depDate, arrDate) & "', 'mm/dd/yy hh:nn')) " & _
''            "Or (Format(nz(depDate, arrDate),'mm/dd/yy hh:nn') Between Format('" & arrDate & "', 'mm/dd/yy hh:nn') And Format('" & Nz(depDate, arrDate) & "', 'mm/dd/yy hh:nn')))")
''        Set rsDUPE = CurrentDb.OpenRecordset("SELECT * " & _
''        "FROM tblPPR " & _
''        "WHERE tblPPR.Status <> 'Cancelled' and tblPPR.archive = False and tblPPR.PPR <> '" & PPR & "' " & _
''        "AND (tblPPR.Callsign = '" & Callsign & "' Or (nz(tblPPR.Tail) <> '' And tblPPR.Tail = '" & Nz(Tail) & "')   ) " & _
''        "AND ((tblPPR.arrdate Between dateserial(year('" & arrDate & "'),month('" & arrDate & "'),day('" & arrDate & "')) And dateserial(year('" & Nz(depDate, arrDate) & "'),month('" & Nz(depDate, arrDate) & "'),day('" & Nz(depDate, arrDate) & "'))) " & _
''            "Or (nz(depDate, arrDate) Between dateserial(year('" & arrDate & "'),month('" & arrDate & "'),day('" & arrDate & "')) And dateserial(year('" & Nz(depDate, arrDate) & "'),month('" & Nz(depDate, arrDate) & "'),day('" & Nz(depDate, arrDate) & "'))))")
'        Set rsDUPE = CurrentDb.OpenRecordset("SELECT * " & _
'        "FROM tblPPR " & _
'        "WHERE tblPPR.Status <> 'Cancelled' and tblPPR.archive = False and tblPPR.PPR <> '" & PPR & "' " & _
'        "AND (tblPPR.Callsign = '" & Callsign & "' Or (nz(tblPPR.Tail) <> '' And tblPPR.Tail = '" & Nz(Tail) & "')   ) " & _
'        "AND (" & _
'            "((tblPPR.arrdate Between #" & arrDate & "# And #" & Nz(depDate, arrDate) & "#) " & _
'                "Or (nz(depDate, arrDate) Between #" & arrDate & "# And #" & Nz(depDate, arrDate) & "#)) " & _
'            "OR ((#" & arrDate & "# Between arrdate And Nz(depDate, arrDate)) " & _
'                "Or (#" & Nz(depDate, arrDate) & "# Between arrDate And Nz(depDate, arrDate)))" & _
'        ")")
'
'
'
'        With rsDUPE
'            If Not .EOF Then
'                .MoveLast
'                .MoveFirst
'
'                If .RecordCount <> 0 And !ID <> Me.ID Then 'If ANY duplicate is found, show to user
'                    If MsgBox("This PPR is conflicting with an existing PPR:" & vbCrLf & vbCrLf & _
'                    "PPR: " & !PPR & vbCrLf & _
'                    "Callsign: " & !Callsign & "   Type: " & !Type & "   Tail: " & !Tail & vbCrLf & _
'                    "Arriving on " & Format(!arrDate, "dd-mmm-yy hh:nn") & "(L)   From: " & !depPoint & vbCrLf & _
'                    "Departing on " & Format(!depDate, "dd-mmm-yy hh:nn") & "(L)   To: " & !Destination & vbCrLf & _
'                    "POC: " & !pocName & " | " & !ctcInfo & vbCrLf & vbCrLf & _
'                    "Overwrite?", vbQuestion + vbYesNo, "New PPR") = vbYes Then
'
'                        'Overwrite
'''                        If Not IsNull(!spID) Then
'''                            'Delete the duplicate if it is on SharePoint
'''                            sending.Caption = "Removing from SharePoint..."
'''                            DoEvents
'''                            Dim db As DAO.Database
'''                            Set db = CurrentDb
'''                            With db
'''                                .Execute "DELETE FROM [Mission Tracker] WHERE ID = " & rsDUPE!spID
'''
'''                                If .RecordsAffected < 2 Then
'''                                    sending.Caption = "Success!"
'''                                Else
'''                                    log "More than one SP record was affected. This was not supposed to happen. Let's continue anyway!", Me.Name & ".savePPR", "WARN"
'''                                End If
'''                                DoEvents
'''                            End With
'''                        End If
'                        .delete 'Delete the duplicate from the database
'                        'Contine saving the current PPR
'
'                    Else
'                        'Do not overwrite the duplicate
'                        If MsgBox("View duplicate?", vbQuestion + vbYesNo, "PPR") = vbYes Then
'                            'If requested, show user the duplicate
'                            DoCmd.Close acForm, "new_ppr"
'                            DoCmd.OpenForm "new_ppr", , , "PPR = '" & !PPR & "'", acFormEdit, acDialog
'                            Exit Function
'                        Else 'Otherwise, stop everything, because we shouldn't be saving duplicates. :)
'
'                            'GoTo fallback
'
'                            '....but sometimes we have no choice.
'
'                        End If
'                    End If
'                End If
'                .Close
'            End If
'        End With
'
'
'
''        DoCmd.OpenForm "frmDupeView", , , , , acDialog, rsDUPE
''        If dupeCheckFail Then GoTo fallback
'
'bClose = True
'30
'    If NewRecord And Not bDELETE Then
'        If formPPR <> pprGenerator Then 'PPR changed from original generator (above)
'            If MsgBox("Using a manual PPR number is not recommended. Continue?" & vbCrLf & "Select 'No' to reset the PPR number.", vbQuestion + vbYesNo, "PPR") = vbNo Then
'                formPPR = pprGenerator
'                GoTo fallback
'            ElseIf Not IsNull(DLookup("ppr", "tblppr", "ppr = '" & formPPR & "'")) Then
'                MsgBox "This PPR was already issued. Please use a different PPR number."
'                GoTo fallback
'            End If
'        End If
'
'        'MOVED TO onLoad
''        '''''PPR Generator'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''        PPR = Format(arrDate, "yy") & "-" & AMOPS & _
''            Format(DMax("right([ppr],3)", "tblppr", _
''            "left([ppr],2) = '" & Format(arrDate, "yy") & "' AND mid([ppr],instr(1,[ppr],'-') + 1,2) = '" & AMOPS & "'") + 1, "000")
''        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'        If Dirty Then Dirty = False
'        MsgBox "Callsign: " & Callsign & vbCrLf & _
'                "Type: " & Me.Type & vbCrLf & _
'                "Arrival Date: " & Format(arrDate, "dd-mmm-yy @ hh:nn") & vbCrLf & _
'                "Departure Date: " & Format(Nz(depDate), "dd-mmm-yy @ hh:nn") & vbCrLf & _
'                "Spot: " & IIf(Spot = "TBD", "NOT ASSIGNED!", Spot) & vbCrLf & _
'                vbCrLf & _
'                "The PPR number for this mission is: " & formPPR _
'                , vbInformation, "New PPR"
'    End If
'
'    'If approach Then Status = "Approach"
'
'
'40
'    If Not bDELETE Then
'        Set rsFP = CurrentDb.OpenRecordset("qInbound")
'        Dim inout As Boolean
'inoutsearch:
'        With rsFP: Do While Not .EOF
'            If !Tail = Tail And Nz(!PPR) = "" Then
'                If MsgBox("Link PPR to today's flight?" & vbCrLf & vbCrLf & !Callsign & " - " & !Type & " - " & !Tail & " - DOF: " & !depDate, vbQuestion + vbYesNo, "PPR") = vbYes Then
'                    .edit
'                    !PPR = formPPR
'                    If Not inout Then arrID = !ID
'                    !spID = spID
'                    Callsign = !Callsign
'                    .update
'                    Exit Do
'                End If
'            End If
'            .MoveNext
'        Loop: End With
'        If Not inout Then
'            inout = True
'            Set rsFP = CurrentDb.OpenRecordset("qOutbound")
'            GoTo inoutsearch
'        End If
'    Else
'        'archive = True
'        Status = "Cancelled"
'    End If
'
'    '''''SharePoint''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''    If IsNull(spID) And spAccess And Not approach And Not bDELETE Then 'New SP Entry
'''        sending.Caption = "Connecting to SharePoint....."
'''        DoEvents
'''        Set qdf = CurrentDb.QueryDefs("qMissionTracker")
'''        qdf.Parameters("fromDay") = LToZ(arrDate)
'''        Set rs = qdf.OpenRecordset()
'''        sending.Caption = "Sending to SharePoint....."
'''        DoEvents
'''        If Not sperr Then
'''            With rs
'''                .AddNew
'''                ![Mission/Event Type] = "PPR"
'''                ![Start Time] = Now()
'''                ![End Time] = Now()
'''                ![PPR #] = formPPR
'''                ![Call Sign] = Callsign
'''                ![Aircraft Type] = Me.Type
'''                ![Tail Number] = Me.Tail
'''                ![Current ICAO] = depPoint
'''                ![ETA (Z)] = LToZ(arrDate)
'''                ![Next ICAO] = pprDestination
'''                ![ETD (Z)] = LToZ(Nz(depDate))
'''                ![DV Code] = dvCode
'''                If Status = "Pending" And Not Remarks Like "*PENDING APPROVAL*" Then
'''                    Remarks = "*PENDING APPROVAL*" & vbCrLf & Remarks
'''                End If
'''                .Fields("Details") = Remarks
'''                !Fuel = Fuel
'''                ![Parking Spot/Location] = Spot
'''                ![AM Initials] = AMOPS
'''                ![TA Initials] = TA
'''                .Update
'''                .Bookmark = .LastModified
'''                spID = !ID
'''                .Close
'''                sending.Caption = "Success!"
'''                DoEvents
'''            End With
'''            Set rs = Nothing
'''        End If
'''45
'''    ElseIf Not IsNull(spID) And spAccess Then 'Existing SP Entry
'''        sending.Caption = "Connecting to SharePoint....."
'''        DoEvents
'''        Set qdf = CurrentDb.QueryDefs("qMissionTrackerByID")
'''        qdf.Parameters("mtbid") = spID
'''        Set rs = qdf.OpenRecordset()
'''    If rs.RecordCount <> 0 Then
'''        sending.Caption = "Sending to SharePoint....."
'''        DoEvents
'''        If approach Or bDELETE Then
'''            rs.delete
'''            spID = Null
'''        Else
'''            If Not sperr Then
'''                If Not rs.EOF Then
'''                    Set rsFP = CurrentDb.OpenRecordset("SELECT * FROM tblTraffic WHERE ID = " & Nz(arrID, 0))
'''                    rs.edit
'''                    If Status = "Cancelled" Or Status = "Denied" Then
'''                        rs![Mission/Event Type] = "Cancelled"
'''                    Else
'''                        Select Case rs![Mission/Event Type]
'''                            Case "Cancelled"
'''                                rs![Mission/Event Type] = "PPR"
'''                        End Select
'''                    End If
'''                    'rs![Mission/Event Type] = IIf(Not Status = "Cancelled", rs![Mission/Event Type], "Cancelled")
'''
'''                    With rs
'''                        ![Start Time] = Nz(![Start Time])
'''                        ![End Time] = Nz(![End Time])
'''                        ![PPR #] = formPPR
'''                        ![Call Sign] = Callsign
'''                        ![Aircraft Type] = Me.Type
'''
'''                        If Not ![Tail Number] Like Me.Tail Then
'''                            If Nz(![Tail Number]) = "" Then
'''                                ![Tail Number] = Me.Tail
'''                            ElseIf MsgBox("Tail number does not match SharePoint:" & vbCrLf & "SharePoint Tail: " & Nz(![Tail Number], "None") & vbCrLf & "Your Tail: " & Tail & _
'''                                    vbCrLf & vbCrLf & "Update SharePoint?", vbQuestion + vbYesNo, "Flight Plan") = vbYes Then
'''                               ![Tail Number] = Me.Tail
'''                            Else
'''                                Me.Tail = ![Tail Number]
'''                            End If
'''                        End If
'''
'''                        ![Current ICAO] = depPoint
'''                        If Not rsFP.EOF Then
'''                            ![ETA (Z)] = LToZ(Nz(rsFP!arrDate, arrDate))
'''                        Else
'''                            ![ETA (Z)] = LToZ(arrDate)
'''                        End If
'''                        ![Next ICAO] = pprDestination
'''                        ![ETD (Z)] = LToZ(Nz(depDate))
'''                        ![DV Code] = dvCode
'''                        If Status = "Pending" And Not Remarks Like "*PENDING APPROVAL*" Then
'''                            Remarks = "*PENDING APPROVAL*" & vbCrLf & Remarks
'''                        End If
'''                        .Fields("Details") = Remarks
'''                        !Fuel = Fuel
'''                        '![Parking Spot/Location] = Spot
'''
'''                        '''''Parking''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''                        Select Case Nz(Spot)
'''                            Case "", "AMC", "TBD"  'We do not have assignment
'''                                Select Case Nz(![Parking Spot/Location])
'''                                Case "", "AMC", "TBD"
'''                                    If Spot <> Nz(![Parking Spot/Location]) Then ![Parking Spot/Location] = Spot
'''                                Case Else
'''                                    GoTo els
'''                                End Select
'''
'''                                Spot = IIf(Left(Nz(![Parking Spot/Location]), 3) = "HOT", "HC" & Right(Nz(![Parking Spot/Location]), 1), Nz(![Parking Spot/Location]))
'''                                '![Parking Spot/Location] = Spot
'''
'''                            Case Is <> ![Parking Spot/Location] 'We have assignment, but it doesnt match the SharePoint
'''els:
'''                                If Nz(![Parking Spot/Location]) = "" Then
'''
'''                                    ![Parking Spot/Location] = Spot
'''
'''                                ElseIf MsgBox("Parking assignment does not match SharePoint:" & vbCrLf & "SharePoint Spot: " & ![Parking Spot/Location] & vbCrLf & "Your Spot: " & Spot & _
'''                                vbCrLf & vbCrLf & "Update SharePoint?", vbQuestion + vbYesNo, "Flight Plan") = vbYes Then
'''                                    ![Parking Spot/Location] = Spot
'''                                Else
'''                                    Spot = ![Parking Spot/Location]
'''                                End If
'''                        End Select
'''
'''                        Dim tlog As DAO.Recordset
'''                        Set tlog = CurrentDb.OpenRecordset("SELECT * FROM tblTraffic WHERE tblTraffic.ID = " & Nz(arrID, 0))
'''                        With tlog
'''                            If .RecordCount = 1 Then
'''                                .edit
'''                                !Spot = Spot
'''                                .Update
'''                                .Close
'''                            End If
'''                        End With
'''                        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''
'''                        ![AM Initials] = AMOPS
'''                        ![TA Initials] = TA
'''                        .Update
'''                        .Close
'''                        sending.Caption = "Success!"
'''                        DoEvents
'''                    End With
'''                 End If
'''            End If
'''        End If
'''    Else
'''        spID = Null
'''    End If
'''    End If
'    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'50
''    DoCmd.RunCommand acCmdSaveRecord
'    Set RS = Nothing
'    Set qdf = Nothing
'
'
'save_exit:
'    savePPR = True
'    Exit Function
'
'fallback:
'    save.enabled = True
'    approval.enabled = True
'    inout = False
'    resOverride = False
'    bClose = False
'    sending.Visible = False
'    Exit Function
'
'save_err:
'    Select Case err
'    Case 3011
'        sperr = True
'        Resume Next
'    Case 3044
'        If onShareDriveDisconnect Then GoTo 10
'    Case 3851
'        Util.relinkTables
'        MsgBox "Something went wrong. Try again in 5 seconds. (3851)", vbCritical, "PPR"
'    Case Else
'        Util.ErrHandler err, Error$, Me.Name & ".savePPR"
'    End Select
'GoTo fallback
'Resume Next
'End Function
'

Private Sub unArchive_Click()
bClose = True
'archive = False
If Dirty Then Dirty = False
DoCmd.OpenForm "new_ppr", , , "PPR = '" & formPPR & "'", acFormEdit
MsgBox "Removed PPR from archive", vbInformation, "PPR"
End Sub


