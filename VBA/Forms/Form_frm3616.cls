VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frm3616"
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

Private Sub addEntry_Click()
On Error GoTo errtrap
Dim ex As Boolean
Dim RS As DAO.Recordset
Dim oi As String: oi = getOpInitials
    For Each ctl In Controls
        If TypeOf ctl Is TextBox Then
            If Nz(ctl.Value) = "" Then
                ctl.BorderColor = RGB(255, 0, 0)
                ex = True
            Else
                ctl.BorderColor = RGB(166, 166, 166)
            End If
        End If
    Next
    If ex Then Exit Sub
    
    If Left(Right(entry, 3), 1) = "…" Then entry = Left(entry, Len(entry) - 3)
    If Left(Right(entry, 5), 3) = "..." Then
        oi = IIf(Right(entry, 2) = getOpInitials, getOpInitials, Right(entry, 2) & "*")
        entry = Left(entry, Len(entry) - 5)
    End If
    If Right(entry, 1) = "." Then entry = Left(entry, Len(entry) - 1)
    entry = UCase(entry)

'    Set rs = CurrentDb.OpenRecordset(logView.RowSource)
'    With rs
'    Do While Not .EOF
'        If !entryTime = LToZ(Format(edate, "dd-mmm-yy") & " " & ZToL(Format(Left(entryTime, 2) & "." & Right(entryTime, 2), "hh:nn"), True)) Then
'            If MsgBox("The following entry will be replaced:" & vbCrLf & Me.entryTime & ": " & !entry & vbCrLf & vbCrLf & "Replace?", vbQuestion + vbYesNo, "Events Log") = vbYes Then
'                .edit
'                !entry = UCase(Me.entry)
'                .update
'                GoTo sexit
'            Else
'                Exit Sub
'            End If
'        End If
'        .MoveNext
'    Loop
    
    entryTime = Format(ZToL(TimeSerial(Left(entryTime, 2), Right(entryTime, 2), 0), True), "hhnn")
    
'    .AddNew
'    !entryTime = LToZ(Format(edate, "dd-mmm-yy") & " " & Format(Left(entryTime, 2) & "." & Right(entryTime, 2), "hh:nn"))
'    If !entryTime < LToZ(DLookup("shiftstart", "tblshiftmanager", "shiftid = " & shiftID)) Then !entryTime = DateAdd("d", 1, !entryTime)
'    'edate & " " & Left(entryTime, 2) & "." & Right(entryTime, 2)
'    !shiftID = shiftID
'    !Initials = DLookup("opinitials", "tbluserauth", "username = '" & Environ$("username") & "'")
'    !entry = entry
'    .update
    Dim midFix As Date: midFix = edate
    If LToZ(DateSerial(Year(edate), Month(edate), Day(edate)) & " " & TimeSerial(Left(entryTime, 2), Right(entryTime, 2), 1)) _
            < DLookup("shiftstart", "tblshiftmanager", "shiftid = " & shiftID) Then
        midFix = DateAdd("d", 1, edate)
    End If


    UtilForm3616.newEntry shiftID, LToZ(DateSerial(Year(midFix), Month(midFix), Day(midFix)) & " " & TimeSerial(Left(entryTime, 2), Right(entryTime, 2), 0)), entry, oi

    
'End With

sexit:
    entryTime = ""
    entry = ""
    logView.Requery
    Exit Sub
errtrap:
    MsgBox Error$
    Resume Next
End Sub

Private Sub checkInsp_Click()
    DoCmd.OpenForm "frmInspections", , , , , acDialog, shiftID
logView.Requery
End Sub

Private Sub btnNewWx_Click()
updateWWAs
End Sub

Private Sub btnTemplateSettings_Click()
'Open some edit for thingy
End Sub

Private Sub cboTemplate_AfterUpdate()
entryTime = IIf(Nz(entryTime) = "", Format(LToZ(Now), "hhnn"), entryTime)
entry.SetFocus
entry = DLookup("verbiage", "tbl3616template", "ID = " & cboTemplate)
cboTemplate = ""
End Sub

Private Sub cboTemplate_Change()
'entryTime = Format(LToZ(Now), "hhnn")
'entry.SetFocus
'entry = DLookup("verbiage", "tbl3616template", "ID = " & cboTemplate)
'cboTemplate = ""
End Sub

Private Sub clear_Click()
entryTime = ""
entry = ""
logView = ""
End Sub

Private Sub closeShift_Click()
Dim RS As DAO.Recordset
Dim rss As DAO.Recordset
Set RS = CurrentDb.OpenRecordset(logView.RowSource)
Set rss = CurrentDb.OpenRecordset("SELECT * FROM tblShiftManager WHERE shiftID = " & shiftID)
Dim closeSt As Boolean

    Select Case MsgBox("Close Shift?", vbQuestion + vbYesNoCancel, "Events Log")
    Case vbYes
        closeSt = True
    Case vbCancel
        Exit Sub
    End Select
    With RS
'    If closeSt Then
'        .AddNew
'        !shiftID = Me.shiftID
'        !entryTime = LToZ(DLookup("shiftEnd", "tblShiftManager", "shiftID = " & shiftID))
'        !Initials = DLookup("opInitials", "tblUserAuth", "username = '" & Environ$("username") & "'")
'
'        With rss
'            rs!entry = UCase(!superLead) & _
'            IIf(Not IsNull(!amoc1), IIf(Not IsNull(!amoc2), ", ", " AND ") & UCase(Nz(!amoc1)), "") & _
'            IIf(Not IsNull(!amoc2), IIf(Not IsNull(!amoc3), ", ", " AND ") & UCase(Nz(!amoc2)), "") & _
'            IIf(Not IsNull(!amoc3), IIf(Not IsNull(!amoc4), ", ", " AND ") & UCase(Nz(!amoc3)), "") & _
'            IIf(Not IsNull(!amoc4), " AND " & UCase(Nz(!amoc4)), "") & " OFF DUTY. " & _
'            IIf(!shiftTitle = "SWINGSHIFT", "AERODROME CLOSED. ", "") & "LOG CLOSED"
'        End With
'        .update
'        .close
'    End If
    End With
    With rss
        .edit
        !closed = True
        .update
        .Close
    End With
    
    If MsgBox("Log closed successfully. Generate PDF?", vbQuestion + vbYesNo, "Events Log") = vbYes Then
        DoCmd.OpenReport "new3616", acViewReport, , "format(shiftstart,'dd-mmm-yy') = format('" & edate & "','dd-mmm-yy')"
    Else
        DoCmd.Close acForm, "frm3616", acSaveYes
    End If
    Set RS = Nothing
    
End Sub

Private Sub Ctl483_Click()
    DoCmd.OpenForm "frm483CompCheck", , , , acFormAdd, , shiftID
    logView.Requery
End Sub

Private Sub delete_Click()
Dim RS As DAO.Recordset

    If Nz(logView) = "" Then
        MsgBox "Select an entry to delete.", vbInformation, "Events Log"
        Exit Sub
    End If
    
    Set RS = CurrentDb.OpenRecordset("SELECT * FROM tbl3616 WHERE ID = " & logView)
    With RS
    If Not .EOF Then
        If MsgBox("Delete entry?", vbQuestion + vbYesNo, "Events Log") = vbYes Then .delete
    Else
        MsgBox "Select an entry to delete.", vbInformation, "Events Log"
    End If
    .Close
    End With
    
    logView.Requery
    entryTime = ""
    entry = ""
    
End Sub

Private Sub Detail_Click()
For Each ctl In Controls
    If TypeOf ctl Is ListBox And Not ctl.Name = "dateSelect" Then
        ctl.Value = ""
    End If
Next
End Sub

Private Sub entry_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
    Case VBA.vbKeyReturn 'Enter key
        bClose = True
        If Dirty Then Dirty = False
        addEntry.SetFocus
        addEntry_Click
        entryTime = ""
        entryTime.SetFocus
    
'    Case VBA.vbKeyF8
'        Dim s As Integer
'        If InStr(1, entry.Text, "_") = 0 Then Exit Sub
'        entry.SelStart = InStr(1, entry.Text, "_") - 1
'
'        s = 1
'        Do While (Right(Mid(entry.Text, entry.SelStart + 1, s + 1), 1) = "_") And (s + entry.SelStart < Len(entry.Text))
'            s = s + 1
'        Loop
'    entry.SelLength = s + 1
'    DoEvents
    End Select
End Sub

Private Sub entry_KeyUp(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
    Case VBA.vbKeyF8
        Dim s As Integer
        If InStr(1, entry.Text, "_") = 0 Then Exit Sub
        entry.SelStart = InStr(1, entry.Text, "_") - 1

        s = 1
        Do While (Right(Mid(entry.Text, entry.SelStart + 1, s + 1), 1) = "_") And (s + entry.SelStart < Len(entry.Text))
            s = s + 1
        Loop
    entry.SelLength = s
    DoEvents
    End Select
End Sub

Private Sub entryTime_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = VBA.vbKeyN Then '"N"
    entryTime = Format(LToZ(Now), "hhnn")
    entry.SetFocus
End If
End Sub

Private Sub Form_Load()
On Error Resume Next
shiftID = Nz(Me.OpenArgs, 0)
If Parent.Name = "frmShiftBoard" Then shiftID = Nz(Parent.shiftID, Parent.OpenArgs)
If shiftID = 0 Then
    If Parent.Parent.getShiftID = 0 Then
        shiftID = Nz(OpenArgs, Parent.Parent.getShiftID)
    Else
        shiftID = Nz(OpenArgs, Parent.Parent.getShiftID)
    End If
End If
Dim RS As DAO.Recordset
logView.RowSource = "SELECT ID, entryTime, entry, initials, shiftID FROM tbl3616 WHERE shiftID = " & shiftID & " ORDER BY entryTime;"

edate = Format(DLookup("shiftStart", "tblShiftManager", "shiftID = " & shiftID), "dd-mmm-yy")

btnTemplateSettings = (DLookup("authlevel", "tbluserauth", "username = '" & getUSN & "'") > 5)

'closeShift.Visible = Not DLookup("closed", "tblShiftManager", "shiftid=" & shiftID)
End Sub

Private Function populate(ByVal sid As Integer)
    logView.RowSource = "SELECT ID, entryTime, entry, initials, shiftID FROM tbl3616 WHERE shiftID = " & sid & " ORDER BY entryTime;"
End Function

Private Sub newlog_Click()
    DoCmd.OpenForm "frmShiftSelect", , , , acFormAdd
End Sub

Private Sub report_Click()
    DoCmd.OpenReport "new3616", acViewPreview, , "Format(ZToL([entryTime]),'mm/dd/yy') = Format('" & dateSelect.Column(1) & "', 'mm/dd/yy') AND [shiftID] = " & dateSelect.Column(0)

End Sub

'Private Sub Form_Open(cancel As Integer)
''Dim isClosed As String
''shiftID = DLookup("lastShift", "tblUserAuth", "username = '" & Environ$("username") & "'")
''isClosed = Nz(DLookup("closed", "tblShiftManager", "shiftID = " & Nz(shiftID, 0)), True)
''
''    If isClosed Then
''        DoCmd.OpenForm ("frmShiftSelect")
''        Cancel = True
''        Exit Sub
''    End If
''    If IsNull(OpenArgs) Then
''        MsgBox "This form was opened illegally!" & vbCrLf & "OpenArgs: " & Nz(OpenArgs, "NULL!"), vbCritical, "AeroStat"
''        'Cancel = True
''    Else
''        logView.RowSource = "SELECT ID, entryTime, entry, initials, shiftID FROM tbl3616 WHERE shiftID = " & Me.OpenArgs & " ORDER BY entryTime;"
''    End If
'
'End Sub

'Private Sub submit_Click()
'Dim rs As DAO.Recordset
'Dim inShift As Boolean
'If IsNull(nEntry) Then Exit Sub
'If IsNull(etime) Then
'    etime.BorderColor = RGB(255, 0, 0)
'    Exit Sub
'Else
'    etime.BorderColor = RGB(166, 166, 166)
'End If
''check if user is part of selected shift
''check for duplicate times
''get entry time using date from selected shift + entryTime
''???
''Profit.
'
'    If IsNull(dateSelect) Then
'        MsgBox "Select a shift log to edit.", vbInformation, "Events Log"
'        Exit Sub
'    End If
'
'    Set rs = CurrentDb.OpenRecordset("SELECT * FROM tblShiftManager WHERE shiftID = " & dateSelect.Column(0))
'    With rs
'    If .RecordCount = 1 Then
'        For Each fld In .Fields
'
'            If Right(fld, 2) = DLookup("opinitials", "tbluserauth", "username = '" & Environ$("username") & "'") Then
'                inShift = True
'                Exit For
'            End If
'        Next
'
'        If Not inShift Then
'
'            If MsgBox("You are not logged into a shift. Create a new one?", vbQuestion + vbYesNo, "AeroStat") = vbYes Then
'                DoCmd.OpenForm "frmShiftSelect", , , , acFormAdd
'            Else
'                Exit Sub
'            End If
'        End If
'    End If
'    .close
'    End With
'
'    Set rs = CurrentDb.OpenRecordset("SELECT * FROM tbl3616 WHERE Format([entryTime],'mm/dd/yy') = Format('" & dateSelect.Column(1) & "','mm/dd/yy') AND shiftID = " & dateSelect)
'    With rs
'    Do While Not .EOF
'
'        If Format(!entryTime, "hhnn") = etime Then
'            MsgBox "An entry for this time already exists. Either offset by 1 minute or change the original entry.", vbInformation, "Event Log"
'            Exit Sub
'        End If
'        .MoveNext
'    Loop
'    .close
'    End With
'
'    etime = ZToL(Format(Left(etime, 2) & "." & Right(etime, 2), "hh:nn"), True)
'
'    Set rs = CurrentDb.OpenRecordset("tbl3616")
'    With rs
'        .AddNew
'        !shiftID = dateSelect
'        !entryTime = Format(dateSelect.Column(1), "dd-mmm-yy") & " " & Format(Left(etime, 2) & "." & Right(etime, 2), "hh:nn")
'        !entryTime = LToZ(!entryTime)
'        !entry = IIf(Not InStr(1, Right(nEntry, 5), "...") = 0, Left(nEntry, Len(nEntry) - 5), nEntry)
'        !Initials = DLookup("opinitials", "tblUserAuth", "username = '" & Environ$("username") & "'")
'        .Update
'        .close
'    End With
'    populate
'    etime = ""
'    nEntry = ""
'
'End Sub

Private Sub logView_DblClick(cancel As Integer)
    If Not IsNull(logView.Column(0)) Then
        entryTime = Format(logView.Column(1), "hhnn")
        entry = DLookup("entry", "tbl3616", "ID = " & logView.Column(0))
    End If
End Sub

Private Sub newNotam_Click()
    DoCmd.OpenForm "frmNOTAMParse", , , , , acDialog, shiftID
    logView.Requery
End Sub

Private Sub preview_Click()
    DoCmd.OpenReport "new3616", acViewReport, , , , edate
End Sub

Private Sub timeNow_Click()
    entryTime = Format(LToZ(Now), "hhnn")
    entry.SetFocus
End Sub
