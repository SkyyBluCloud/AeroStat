VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_check_in_out"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'   AeroStat: Airfield Management(TM)
'
'
'    Copyright 2017-2020 Shane J. DeSouza(SkyyBluCloud)
'
'   Licensed under the Apache License, Version 2.0 (the "License");
'   you may not use this file except in compliance with the License.
'   You may obtain a copy of the License at
'
'       http://www.apache.org/licenses/LICENSE-2.0
'
'   Unless required by applicable law or agreed to in writing, software
'   distributed under the License is distributed on an "AS IS" BASIS,
'   WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
'   See the License for the specific language governing permissions and
'   limitations under the License.

Option Compare Database
Dim bClose As Boolean

Private Sub Form_BeforeUpdate(cancel As Integer)
    If Not bClose Then
        cancel = True
        Undo
        Exit Sub
    End If
End Sub

Private Sub Form_Error(DataErr As Integer, response As Integer)
    If DataErr = 2169 Then response = True
End Sub

Private Sub Form_Current()
On Error GoTo errtrap
Dim stn As String
bClose = False

stn = DLookup("data", "tblSettings", "key = 'station'")
reload:
    If Status = "Pending" Then
        If depPoint = Destination Or Not Destination = stn Then
            TWR_out.Visible = True
            CP_out.Visible = True
            
        ElseIf Not depPoint = stn Then
            lblor.Visible = True
            ctlETA.Visible = True
        End If
        
        ctlATD.Visible = True
        ctlATD = Format(LToZ(Now), "hhnn")
        ctlATD.SetFocus
        DOF.Visible = True
        
    ElseIf Status = "Enroute" Then
        If depPoint = Destination Or Not depPoint = stn Then
            ctlATA.Visible = True
            ctlATA = Format(LToZ(Now), "hhnn")
            ctlATA.SetFocus
            TWR_in.Visible = True
            CP_in.Visible = True
            edit.Visible = False
            undoDep.Visible = True
            Spot.Visible = True
            lblSpot.Visible = True
            If Not IsNull(ctlETA) Then undoDep.Caption = "Reset ATD"
            
        ElseIf depPoint = stn Then
            If MsgBox("This will reset the times. Continue?", vbQuestion + vbYesNo, "Traffic Log") = vbYes Then
                ctlATA = ""
                ctlATD = ""
                ATA = ""
                ATD = ""
                Status = "Pending"
                
                bClose = True
                If Dirty Then Dirty = False
                GoTo reload
                bClose = False
            Else
                DoCmd.Close acForm, "check_in_out"
            End If
            
            Exit Sub
            
        End If
        
    ElseIf Status = "Closed" Then
        If MsgBox("This will reset the times. Continue?", vbQuestion + vbYesNo, "Traffic Log") = vbYes Then
            ctlATA = ""
            ctlATD = ""
            ATA = ""
            ATD = ""
            Status = "Pending"
            
            bClose = True
            If Dirty Then Dirty = False
            
        End If
        DoCmd.Close acForm, "check_in_out"
        Exit Sub
    End If
sexit:
    If Not IsNull(ETA) Then ctlETA = Hour(ETA) & ":" & Format(Minute(ETA), "00")
    If Not IsNull(ATA) Then ctlATA = Hour(ATA) & ":" & Format(Minute(ATA), "00")
    If Not IsNull(ATD) Then ctlATD = Hour(ETD) & ":" & Format(Minute(ETD), "00")
    Exit Sub
errtrap:
    ErrHandler err, Error$, Me.Name & ".Form_Current"
    Resume Next
    
    
    
    
End Sub

Private Sub Form_Close()
'    Dim tl As Variant
'    If CurrentProject.AllForms("frmMain").IsLoaded Then
'        Set tl = Forms!frmMain!view
'    Else
'        Exit Sub
'    End If
'
'    For Each ctl In tl.Controls
'        If TypeOf ctl Is ListBox Then ctl.Requery
'    Next
End Sub

Private Sub ctlATD_BeforeUpdate(cancel As Integer)
On Error Resume Next
If Not IsNumeric(ctlATD) Or Len(ctlATD) > 4 Then
    ctlATD = ""
    Exit Sub
End If
If Nz(ctlATD) = "" Then Exit Sub

ATD = getTime4Char(ctlATD)
ctlATD = Left(ctlATD, 2) & ":" & Right(ctlATD, 2)
End Sub

Private Sub ctlATA_BeforeUpdate(cancel As Integer)
On Error Resume Next
If Not IsNumeric(ctlATA) Or Len(ctlATA) > 4 Then
    ctlATA = ""
    Exit Sub
End If
If Nz(ctlATA) = "" Then Exit Sub

ATA = getTime4Char(ctlATA)
ctlATA = Left(ctlATA, 2) & ":" & Right(ctlATA, 2)
End Sub

Private Sub ctlETA_BeforeUpdate(cancel As Integer)
On Error Resume Next
If Not IsNumeric(ctlETA) Or Len(ctlETA) > 4 Then
    ctlETA = ""
    Exit Sub
End If

If Nz(ctlETA) = "" Then Exit Sub

ETA = getTime4Char(ctlETA)
ctlETA = Left(ctlETA, 2) & ":" & Right(ctlETA, 2)
End Sub

Private Sub ctlATD_Exit(cancel As Integer)
    If IsDate(ctlATD) Then Exit Sub
    If IsNumeric(ctlATD) And Len(ctlATD) = 3 Then ctlATD = 0 & ctlATD
    ctlATD.BorderColor = RGB(166, 166, 166)
    ATD = getTime4Char(Nz(ctlATD))
    ctlATD = Left(ctlATD, 2) & ":" & Right(ctlATD, 2)
    
    If Not IsDate(ctlATD) Then
        'cancel = True
        ctlATD = ""
        ctlATD.BorderColor = RGB(255, 0, 0)
        'ctlATD.SetFocus
    End If
End Sub

Private Sub ctlATA_Exit(cancel As Integer)
    If IsDate(ctlATA) Or IsNull(ctlATA) Then Exit Sub
    If IsNumeric(ctlATA) And Len(ctlATA) = 3 Then ctlATA = 0 & ctlATA
    ctlATA.BorderColor = RGB(166, 166, 166)
    ATA = getTime4Char(ctlATA)
    ctlATA = Left(ctlATA, 2) & ":" & Right(ctlATA, 2)
    
    If Not IsDate(ctlATA) Then
        'cancel = True
        ctlATA = ""
        ctlATA.BorderColor = RGB(255, 0, 0)
        'ctlATA.SetFocus
    End If
End Sub

Private Sub ctlETA_Exit(cancel As Integer)
    If IsDate(ctlETA) Then Exit Sub
    If IsNumeric(ctlETA) And Len(ctlETA) = 3 Then ctlETA = 0 & ctlETA
    ctlETA.BorderColor = RGB(166, 166, 166)
    ETA = getTime4Char(ctlETA)
    ctlETA = Left(ctlETA, 2) & ":" & Right(ctlETA, 2)
    
    If Not IsDate(ctlETA) Then
        'cancel = True
        ctlETA = ""
        ctlETA.BorderColor = RGB(255, 0, 0)
        'ctlETA.SetFocus
    End If
End Sub

Private Sub cancel_Click()
On Error GoTo close_Click_Err
    On Error Resume Next
    DoCmd.Close acForm, "check_in_out"

close_Click_Exit:
    Exit Sub

close_Click_Err:
    MsgBox Error$
    Resume close_Click_Exit
End Sub

Private Sub edit_Click()
On Error GoTo err
    bClose = True
    DoCmd.OpenForm "quick_input", , , "ID=" & ID, acFormEdit
    DoCmd.Close acForm, "check_in_out"
    
err: If err.Number = 2501 Then Resume Next
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
    If KeyCode = 13 And Shift = acShiftMask Then
        submit.SetFocus
        submit_Click
    End If
End Sub

Private Sub submit_Click()
On Error GoTo save_Click_Err
For Each ctl In Controls
    If ctl.Tag = "btn" Then ctl.Enabled = False
Next
Dim RS As DAO.Recordset
Dim qdf As DAO.QueryDef
'Dim rstCount As DAO.Recordset
Dim rstSP As DAO.Recordset
Dim stn As String
Dim spfound As Boolean
Dim spAccess As Boolean
Dim tz As Double
bClose = True
spAccess = DLookup("spAccess", "tblUserAuth", "username = '" & Util.getUser & "'")
If spAccess Then spAccess = sp
tz = DLookup("data", "tblSettings", "key = ""timezone""")
stn = DLookup("data", "tblSettings", "key = 'station'")
sending.Visible = True
sp.Visible = False
DoEvents
    
    If ATD = "0:00" Then ATD = "0:01"
    If ETA = "0:00" Then ETA = "0:01"
    If ATA = "0:00" Then ATA = "0:01"
    
    ATD = Format(ATD, "hh.nn")
    ATA = Format(ATA, "hh.nn")
    
    If Not IsNull(ETA) Then ATD = ""
    
    Dim st As Boolean
    For Each ctl In Controls
        If ctl.Tag = "nullck" Then
            If IsNull(ctl.Value) Then
                ctl.BorderColor = RGB(255, 0, 0)
                st = True
            Else
                ctl.BorderColor = RGB(166, 166, 166)
            End If
        End If
    Next
    If st Then Exit Sub
    
    For Each ctl In Controls
        If Left(ctl.Name, 3) = "TWR" Or Left(ctl.Name, 2) = "CP" Then ctl.Value = UCase(ctl.Value)
    Next
    
    '''''Traffic Count''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'Open the appropriate traffic count record for the day
    Dim arrordep As String
    If Destination = stn Then
        'arrDate = Format(ZToL(DOF + (Nz(ATD, ETD) + ETE)), "dd-mmm-yy") & " " & Format(ZToL(IIf(IsNull(ATD), IIf(Not ETA = 0, ETA, ETD + ETE), ATD + ETE)), "hh:nn")
        arrDate = Format(DOF + ZToL(Nz(ATD, ETD) + ETE), "dd-mmm-yy") & " " & Format(ZToL(IIf(IsNull(ATD), IIf(Not ETA = 0, ETA, ETD + ETE), ATD + ETE)), "hh:nn")
        arrordep = arrDate
    End If
    If depPoint = stn Then
        depDate = DOF + ZToL(Nz(ATD, ETD))
        If arrordep = "" Then arrordep = depDate
    End If
    arrordep = Format(arrordep, "dd-mmm-yy")
'
'    Set 'rstCount = CurrentDb.OpenRecordset("SELECT * FROM tblAnalytics WHERE Format(tblAnalytics.countDay,'dd-mmm-yy') = '" & arrordep & "'")
'    If 'rstCount.RecordCount = 0 Then
'        'rstCount.AddNew
'        'rstCount!countDay = arrordep
''        'rstCount.Update
''        'rstCount.Requery
'    Else
'        'rstCount.edit
'    End If
    
'    'Set ICAO name to be recognized by counter
'    If fieldExists(ICAOToCount(Me.acType, Me.isMil) & "arr", "tblAnalytics") Or fieldExists(ICAOToCount(Me.acType, Me.isMil) & "dep", "tblAnalytics") Then
'        countType = ICAOToCount(Me.acType, Me.isMil)
'    Else
'        Select Case isMil
'            Case 0 'US Mil
'                countType = IIf((Mid(Me.acType, 1, 1) = "H") Or (Mid(Me.acType, 2, 1) = "H"), "otherMilHeli", "otherMil")
'
'            Case 1 'Commercial
'                If Mid(Me.Tail, 1, 1) = "N" Then
'                    countType = IIf((Mid(Me.acType, 1, 1) = "H") Or (Mid(Me.acType, 2, 1) = "H"), "otherUSHeli", "otherUSFixed")
'                Else
'                    countType = IIf((Mid(Me.acType, 1, 1) = "H") Or (Mid(Me.acType, 2, 1) = "H"), "otherHeli", "otherFixed")
'                End If
'
'            Case 2 'JSDF
'                countType = "otherJSDF"
'
'        End Select
'    End If
    
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    '''''Initialize Share Point Record''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    If Not Me.acType = "C172" And spAccess Then
        sending.Caption = "Conecting to SharePoint...."
        DoEvents
        
        If IsNull(spID) Then
            Set qdf = CurrentDb.QueryDefs("qMissionTracker")
            qdf.Parameters("fromDay") = Format(arrordep, "mm/dd/yy")
            Set rstSP = qdf.OpenRecordset()
        Else
            Set qdf = CurrentDb.QueryDefs("qMissionTrackerByID")
            qdf.Parameters("mtbid") = spID
            Set rstSP = qdf.OpenRecordset()
        End If
        
        Dim csnospace As String
        Do While Not rstSP.EOF
            If Not IsNull(spID) Then
                spfound = True
                rstSP.edit
                sending.Caption = "Sending to SharePoint...."
                DoEvents
                If IsNull(rstSP![Start Now]) Then rstSP![Start Now] = Date
                If IsNull(rstSP![End Now]) Then rstSP![End Now] = Date
                Exit Do
            End If
            csnospace = Replace(Nz(rstSP![Call Sign]), " ", "")
            If csnospace = Callsign Or spID = rstSP!ID Then
                sending.Caption = "Sending to SharePoint...."
                DoEvents
                spID = rstSP!ID
                rstSP.edit
                If IsNull(rstSP![Start Now]) Then rstSP![Start Now] = Date
                If IsNull(rstSP![End Now]) Then rstSP![End Now] = Date
                spfound = True
                Exit Do
            End If
            rstSP.MoveNext
        Loop
    End If
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    '''''Update Traffic Log/Count'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ''rstCount.edit
    'If plane arrives ANYWHERE
    If (Not IsNull(ATD) Or Not IsNull(ETA)) And Not IsNull(ATA) Then
        Status = "Closed"
        
        'If arrived here
        If Destination = stn Then
            If Not Stationed Then
                If Not findParentByTail(Nz(Tail), 2) = 0 Then
                    Set RS = CurrentDb.OpenRecordset("SELECT * FROM tblTraffic WHERE ID = " & findParentByTail(Tail, 2))
                    With RS
                    If Not .EOF Then
                        .edit
                        !Stationed = True
                        .update
                        .Close
                    End If
                    End With
                Else
                    Stationed = True
                End If
            End If
            ''rstCount.Fields(countType & "arr") = 'rstCount.Fields(countType & "arr") + Me.Number
            'arrDate = Format(ZToL(DOF & " " & Nz(ATD, ETD)), "mm/dd/yy") & " " & Format(ZToL(ATA), "hh:nn")
            'arrDate = Format(DateAdd("n", (60 * Format(ETE, "h")) + Format(ETE, "nn"), ZToL(DOF & " " & Nz(ATD, ETD))), "dd-mmm-yy") & " " & Format(ZToL(ATA), "hh.nn")
            If spfound Then
                rstSP![ATA (Z)] = LToZ(arrDate)
                rstSP![Tower Initials] = TWR_in
                'If IsNull(rstSP![Parking Spot/Location]) Then rstSP![Parking Spot/Location] = Spot
            '''''Parking''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
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
                CurrentDb.Execute "UPDATE tblTraffic SET Spot = '" & Spot & "' WHERE Status Not Like 'C*' AND PPR = '" & PPR & "'"
            End If
        End If
    
    'If departed ANYWHERE
    ElseIf Not IsNull(ATD) Or Not IsNull(ETA) Then
        ATA = ""
        'If Not DateDiff("d", depDate, Now) = 0 Then
'        If (tz < 0 And DateValue(ZToL(DOF + ATD)) < DateValue(ZToL(DOF + ETD))) Or (tz > 0 And DateValue(ZToL(DOF + ATD)) > DateValue(ZToL(DOF + ETD))) Then
        If Abs(DateDiff("h", ATD, ETD)) > 6 And _
            (tz < 0 And DateValue(ZToL(DOF + ATD)) < DateValue(ZToL(DOF + ETD)) Or _
            tz > 0 And DateValue(ZToL(DOF + ATD)) > DateValue(ZToL(DOF + ETD))) Then
            
            'DOF = DOF + 1
            If MsgBox("Confirm DOF:" & vbCrLf & Format(DOF, "dd-mmm-yy") & vbCrLf & vbCrLf & "Is this correct?" & vbCrLf & "(You may have to change the DOF if the departure was past 0000Z)", vbQuestion + vbYesNo, "AeroStat") = vbNo Then
                GoTo fb
            End If
            depDate = DOF + ZToL(Nz(ATD, ETD))
        End If
        
        Status = "Enroute"
        'arrDate = Format(ZToL(DOF + (Nz(ATD, ETD) + ETE)), "mm/dd/yy") & " " & Format(ZToL(cETA, True), "hh:nn")
        arrDate = DateAdd("n", (60 * Format(ETE, "h")) + Format(ETE, "nn"), DOF + ZToL(Nz(ATD, ETD)))
        'if departed here
        If depPoint = stn Then
            
            If Stationed Then
                Stationed = False
            ElseIf Not IsNull(ParentID) Then
                Set RS = CurrentDb.OpenRecordset("SELECT * FROM tblTraffic WHERE ID = " & ParentID)
                With RS
                If Not .EOF Then
                    .edit
                    !Stationed = False
                    .update
                    .Close
                End If
                End With
            End If
            'rstCount.Fields(countType & "dep") = 'rstCount.Fields(countType & "dep") + Me.Number
            'If Not checkParking(Callsign, Nz(Tail)) = 0 Then ParentID = checkParking(Callsign, Nz(Tail))
            'If Destination = stn Then arrDate = DOF + DateAdd("h", DLookup('data','tblSettings','key = ""timezone""'), IIf(Status = "Enroute" Or Status = "Closed", IIf(Status = "Enroute", cETA, ATA), cETA))
            
            
            If spfound Then
                rstSP![ATD (Z)] = LToZ(depDate)
                rstSP![Tower Initials] = TWR_out
            End If
        Else
            If spfound Then
                rstSP![ETA (Z)] = DateAdd("h", -DLookup("data", "tblSettings", "key = ""timezone"""), arrDate)
            End If
        End If
    Else
        Status = "Pending"
    End If
    
    'rstCount.Update
    If spfound Then
        rstSP.update
        sending.Caption = "Success!"
    Else
        sending.Caption = "Not Found"
    End If
    DoEvents
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    '''''Update PPR info''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    If Not Destination = depPoint Then ' Not a local flt
    
        'This is probably redundant...
        'If Destination = stn Then arrDate = DOF + DateAdd("h", DLookup('data','tblSettings','key = ""timezone""'), IIf(Status = "Enroute" Or Status = "Closed", IIf(Status = "Enroute", cETA, ATA), cETA))
        'If depPoint = stn Then depDate = DOF + DateAdd("h", DLookup('data','tblSettings','key = ""timezone""'), IIf(Status = "Enroute", ATD, ETD))

        If Not Nz(PPR) = "" Then
            Set RS = CurrentDb.OpenRecordset("SELECT * FROM tblPPR WHERE tblPPR.PPR = '" & PPR & "'")

            If Not RS.EOF Then
                RS.edit
                'If Destination = stn Then rs!arrDate = arrDate
                If depPoint = stn Then RS!depDate = depDate
                If Destination = stn And Nz(Spot) = "" Then
                    Spot = RS!Spot
                Else
                    RS!Spot = Spot
                End If
                RS.update
'                Set rs = CurrentDb.OpenRecordset("SELECT tblTraffic.ID, tblTraffic.Stationed, tblTraffic.Status FROM tblTraffic WHERE tblTraffic.ID = " & rs!ParentID)
'                rs.edit
'                If rs!Stationed Then rs!Stationed = False
'                rs!Status = IIf(Stationed, Spot, IIf(Not Destination = stn, "Closed", Status))
'                rs.update
            End If
            RS.Close
        End If
    End If
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    'rstCount.close
    Dirty = False
    Set RS = Nothing
    Set qdf = Nothing
    Set rstSP = Nothing

save_Click_Exit:
    'syncTrafficLog ID, "Traffic", False
    FormSyncUtil.syncForm "Traffic Log"

    DoCmd.Close acForm, "check_in_out"
    If Forms!frmMain!view.SourceObject = "Traffic Log" Then
        For Each ctl In Forms!frmMain!view.Controls
            If TypeOf ctl Is ListBox Then ctl.Requery
        Next
        Forms!frmMain!view.Form.osReload
    End If
    Exit Sub
    
fb:
    sending.Caption = "Loading"
    sending.Visible = False
    sp.Visible = True
    bClose = False
    For Each ctl In Controls
        If ctl.Tag = "btn" Then ctl.Enabled = True
    Next
    Exit Sub

save_Click_Err:
    MsgBox Error$
    Resume Next
End Sub

Private Sub undoDep_Click()
Dim RS As DAO.Recordset
Dim reopen As Integer
reopen = ID
    If IsNull(ETA) Then If MsgBox("Cancel DEP message?", vbYesNo, "Undo") = vbNo Then Exit Sub
    bClose = True
    Undo
    ATD = ""
    ETA = ""
    TWR_out = ""
    Status = "Pending"
    DoCmd.RunCommand acCmdSaveRecord
    If Not IsNull(ParentID) Then
        Set RS = CurrentDb.OpenRecordset("SELECT * FROM tblTraffic WHERE tblTraffic.ID = " & ParentID)
        RS.edit
        'rs!Status = "Pending"
        RS.update
    End If

    DoCmd.Close acForm, "check_in_out"
    'DoCmd.OpenForm "check_in_out", , , "ID = " & reopen, acFormEdit
End Sub

