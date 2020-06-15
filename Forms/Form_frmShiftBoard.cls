VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmShiftBoard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Public bClose As Boolean
Dim lstOnShiftChange As New Collection

Private Sub btnReloadNOTAMs_Click()
notamM = DCount("NOTAM", "qNOTAMM", "left([NOTAM],1) = 'M'")
notamL = DCount("NOTAM", "qNOTAML", "left([NOTAM],1) = 'L'")

End Sub

Private Function dirtyCheck()
btnSave = Me.Dirty
End Function

Private Sub btnRemoveAmoc_Click()
If lstOnShift.ItemsSelected.Count = 0 Then Exit Sub

Dim lstOsSel() As String
lstOsSel = getSelection(lstOnShift)

    lstAMOCs.RowSource = lstAMOCs.RowSource & ";" & join(lstOsSel, ";")
    
    Dim authLevel As Integer
    For Each a In lstOsSel
        lstOnShift.RemoveItem (a)
        authLevel = DLookup("authlevel", "tbluserauth", "opInitials = '" & Right(a, 2) & "'")
        lstOnShiftChange.add (a & " " & IIf(authLevel = 8, "(TRAINEE)", "(AMOC)") & " OFF SHIFT")
        
    Next
    
    CurrentDb.Execute "UPDATE tblUserAuth SET onShift = False WHERE ([rank] & ' ' & Left([firstName],1) & '. ' & [lastName] & '/' & [opInitials]) = '" & join(lstOsSel, "' OR ([rank] & ' ' & Left([firstName],1) & '. ' & [lastName] & '/' & [opInitials]) = '") & "'"
    
    Me.superLead.SetFocus
    Me.amoc = lstOnShift.RowSource
    Me.Dirty = True
    btnSave = Dirty
    bClose = False
End Sub

Private Sub btnReview_Click()
DoCmd.OpenForm "frmAfmDash"
End Sub

Private Sub btnSave_Click()
    bClose = True
    
    Dim idx As Integer
    Dim onShiftInitials() As String
    Dim changes() As String
    onShiftInitials = Split(lstOnShift.RowSource, ";")
    
    For Each s In onShiftInitials
        onShiftInitials(idx) = Right(onShiftInitials(idx), 2)
        idx = idx + 1
    Next
    
    If lstOnShiftChange.Count <> 0 Then
        idx = 0
        For Each s In lstOnShiftChange
            ReDim Preserve changes(0 To idx) As String
            changes(idx) = s
            idx = idx + 1
        Next
        
        
        
        If MsgBox("Add log entry?" & vbCrLf & vbCrLf & UCase(join(changes, ". ")), vbQuestion + vbYesNo, "Shift Board") = vbYes Then
            UtilForm3616.newEntry shiftID, LToZ(Now), UCase(join(changes, ". "))
        End If
    End If
        
    
    'Update users on shift
    CurrentDb.Execute "UPDATE tblUserAuth SET lastShift = " & shiftID & ", onShift = True WHERE opInitials = '" & Right(superLead, 2) & "' OR opInitials = '" & _
        join(onShiftInitials, "' OR opInitials = '") & "'"
        
    'Update users who are no longer on the shift
    CurrentDb.Execute "UPDATE tblUserAuth SET onShift = False WHERE lastShift = " & shiftID & " AND (opInitials <> '" & Right(superLead, 2) & "' AND opInitials <> '" & _
        join(onShiftInitials, "' AND opInitials <> '") & "')"
        
    
            
    If Me.Dirty Then Me.Dirty = False
    btnSave = Me.Dirty
    Set lstOnShiftChange = New Collection
End Sub

Private Sub btnStartStop_AfterUpdate()
Select Case btnStartStop
Case True
    btnStartStop.Caption = "Close and Sign"
Case False
    btnStartStop.Enabled = Not Me.closed
    btnStartStop.Caption = "Start Shift"
End Select
End Sub

Private Sub btnStartStop_Click()
Dim recsrc As String
Dim rs As DAO.Recordset
Dim inc As Boolean
bClose = True
recsrc = "SELECT * FROM tblshiftmanager WHERE shiftID = " & shiftID

    Select Case btnStartStop
    Case True 'Shift start
        
        btnReloadNOTAMs_Click
        If Nz(shiftTitle) = "" Then
            shiftTitle.BorderColor = RGB(255, 0, 0)
            btnStartStop = False
            btnStartStop_AfterUpdate
            Exit Sub
        Else
            shiftTitle.BorderColor = RGB(166, 166, 166)
        End If
        
        es = DLookup("shiftID", "tblShiftManager", "Format(shiftStart,'dd-mmm-yy hh:nn') = Format('" & shiftStart & "','dd-mmm-yy hh:nn')")
        If Not IsNull(es) Then
            If MsgBox("This shift was already started. Would you like to view it?", vbQuestion + vbYesNo, "AeroStat") = vbYes Then
                Undo
                DoCmd.OpenForm "frmShiftBOARD", , , "shiftID = " & es, acFormEdit
                Exit Sub
            End If
        End If
        
        '''''Conversion
    '    For Each ctl In Controls
    '        If TypeOf ctl Is TextBox Then
    '            If InStr(1, ctl.Tag, "init") > 0 And Not IsNull(ctl.Value) Then
    '                ctl.Value = getLogName(ctl.Value)
    '            End If
    '        End If
    '    Next
        '''''
        '''''Null check
        For Each ctl In Controls
            If TypeOf ctl Is TextBox Then
                If ctl.Tag Like "*nonull*" Then
                    If Nz(ctl.Value) = "" Then
                        inc = True
                        ctl.BorderColor = RGB(255, 0, 0)
                    Else
                        ctl.BorderColor = RGB(166, 166, 166)
                    End If
                End If
                If ctl.Tag Like "*caps*" Then
                    ctl.Value = UCase(ctl.Value)
                End If
            End If
        Next
        If inc Then
            btnStartStop = False
            btnStartStop_AfterUpdate
            Exit Sub
        End If
        
        Dim msg As String
        Dim idx As Integer
        Dim onShift() As String
        onShift = Split(lstOnShift.RowSource, ";")
        
        For Each s In onShift
            onShift(idx) = Right(onShift(idx), 2)
            idx = idx + 1
        Next
        
        'Update users on shift
        CurrentDb.Execute "UPDATE tblUserAuth SET lastShift = " & shiftID & ", onShift = True WHERE opInitials = '" & Right(superLead, 2) & "' OR opInitials = '" & _
            Nz(join(onShift, "' OR opInitials = '")) & "'"
            
        'Update users who are no longer on the shift
        CurrentDb.Execute "UPDATE tblUserAuth SET onShift = False WHERE lastShift = " & shiftID & " AND opInitials <> '" & Right(superLead, 2) & "' AND opInitials <> '" & _
            join(onShift, "' AND opInitials <> '") & "'"
        
        
        
        If NewRecord Then
            msg = "Shift: " & shiftTitle & " (" & shiftStart & " - " & shiftEnd & ")" & vbCrLf & _
                    "AMOS/AMSL: " & superLead & vbCrLf & _
                    "AMOC: " & Replace(Nz(Me.amoc), ";", ", ")
            
'            For i = 1 To UBound(arrJoined)
'                msg = msg & " " & arrJoined(i)
'                If Not i = UBound(arrJoined) Then msg = msg & ","
'            Next
            
            'Gather all args in the openStatement
            Dim entry As String
            Dim args As New Collection
            entry = DLookup("openstatement", "tblshiftdef", "shifttitle = '" & shiftTitle & "'")
            a = InStr(1, entry, "{")
            Do While a <> 0
                b = InStr(a, entry, "}")
                args.add Mid(entry, a, b - (a - 1))
                a = InStr(b, entry, "{")
            Loop
            
            'Iterate through args and replace with relevant data
            Dim result As String
            Dim authLevel As Integer
            authLevel = DLookup("authlevel", "tbluserauth", "opInitials = '" & Right(superLead, 2) & "'")
            For Each arg In args
            
                Select Case arg
                Case "{ONSHIFT}"
                    'result = superLead & " " & getPos(DLookup("authlevel", "tbluserauth", "opInitials = '" & Right(superLead, 2) & "'"))
                    result = superLead & " " & IIf(authLevel < 6, "(AMOS)", "(AMSL)")
                    For Each c In Split(lstOnShift.RowSource, ";")
                        authLevel = DLookup("authlevel", "tbluserauth", "opInitials = '" & Right(c, 2) & "'")
                        result = result & ", " & c & " " & IIf(authLevel = 8, "(TRAINEE)", "(AMOC)")
                    Next
                    entry = Replace(entry, arg, result)
                    
                Case "{AFLDSTATUS}"
                    If rwy = "AD CLSD" Then
                        result = "AERODROME CLOSED"
                    Else
                        result = "RWY " & rwy & " IN USE. RSC " & rsc & "; BWC " & bwc & "; BARRIERS " & IIf(barrier = "WITHDRAWN", "WITHDRAWN", "CONFIGURED FOR RWY " & barrier)
                    End If
                    entry = Replace(entry, arg, result)
                
                Case "{NOTAMS}"
                    result = notamM & "M, " & notamV & "V, " & notamL & "L"
                    entry = Replace(entry, arg, result)
                
                Case "{NAVAIDS}"
                    result = "NAVAIDS: DASR(" & dasr & "), ILS(" & ils & "), TACAN(" & tacan & ")"
                    entry = Replace(entry, arg, result)
                
                Case "{CRASH}"
                    entry = Replace(entry, arg, "CRASH STATUS " & aarf)
                    
                Case "{OPENCALL}"
                    Dim ocInitials As String
                    Dim opencall() As String
                    
                    If IsNull(DLookup("opencall", "tblshiftdef", "shifttitle = '" & shiftTitle & "'")) Then
                        entry = Replace(entry, arg, "")
                    Else
                        opencall = Split(DLookup("opencall", "tblshiftdef", "shifttitle = '" & shiftTitle & "'"), ",")
                        For Each oc In opencall
                            ocInitials = Nz(ocInitials) & IIf(Nz(ocInitials) = "", "", ", ") & oc & "/" & InputBox("Call/notify and enter initials for " & oc, "Coordination Required")
                        Next
                        result = ocInitials & " NOTIFIED."
                        entry = Replace(entry, arg, result)
                    End If
                End Select
            Next
            
            If Dirty Then Dirty = False
            
            'Send it.
            UtilForm3616.newEntry shiftID, LToZ(shiftStart), entry
        
'            '''''AUTOLOG
'            Set rs = CurrentDb.OpenRecordset("SELECT * FROM tbl3616 WHERE tbl3616.shiftID = " & Me.shiftID & ";")
'            If rs.RecordCount = 0 Then
'            Dim amoclist As String
'            amoclist = join(amoc, ";")
''            al1 = amoclist(1)
''            al2 = Nz(DLookup("authlevel", "tbluserauth", "opInitials = '" & Right(Nz(amoc1), 2) & "'"))
''            al3 = Nz(DLookup("authlevel", "tbluserauth", "opInitials = '" & Right(Nz(amoc2), 2) & "'"))
''            al4 = Nz(DLookup("authlevel", "tbluserauth", "opInitials = '" & Right(Nz(amoc4), 2) & "'"))
'            With rs
'                .AddNew
'                !shiftID = Me.shiftID
'                !entryTime = LToZ(Me.shiftStart)
'                !Initials = DLookup("opInitials", "tblUserAuth", "username = '" & Environ$("username") & "'")
'                'amoc field should store full name instead of initials
'                '!entry = IIf(shiftTitle = "DAYSHIFT", "LOG OPEN. ", "") ...
'                !entry = "LOG OPEN. " & UCase(superLead) & " " & getPos(DLookup("authlevel", "tbluserauth", "opInitials = '" & Right(superLead, 2) & "'")) & _
'                IIf(Not IsNull(amoc1), IIf(Not IsNull(amoc2), ", ", " AND ") & UCase(Nz(amoc1)) & " " & getPos(IIf(al2 = 5 Or al2 = 6, 7, al2)), "") & _
'                IIf(Not IsNull(amoc2), IIf(Not IsNull(amoc3), ", ", " AND ") & UCase(Nz(amoc2)) & " " & getPos(IIf(al3 = 5 Or al3 = 6, 7, al3)), "") & _
'                IIf(Not IsNull(amoc3), IIf(Not IsNull(amoc4), ", ", " AND ") & UCase(Nz(amoc3)) & " " & getPos(IIf(al4 = 5 Or al4 = 6, 7, al4)), "") & _
'                IIf(Not IsNull(amoc4), " AND " & UCase(Nz(amoc4)) & " " & getPos(Nz(DLookup("authlevel", "tbluserauth", "opInitials = '" & Right(Nz(amoc4), 2) & "'"))), "") & " ON DUTY. " & _
'                "NOTAMS: " & notamM & "M, " & notamV & "V, " & notamL & "L. FLT PLANNING ROOM CK CMPLT. " & _
'                IIf(Left(shiftTitle, 10) = "SWINGSHIFT", "RWY " & rwy & " IN USE. RSC " & rsc & ", BWC " & bwc & ". BARRIERS " & IIf(barrier = "WITHDRAWN", "WITHDRAWN", "CONFIGURED FOR RWY " & barrier) & ". ", "") & _
'                shiftTitle & " CHECKLIST INITIATED. MONITORING RAMPNET FOR PROPER PHRASEOLOGY THROUGHOUT SHIFT"
'                'FUTURE: User customization
'                .update
'            End With
'            End If
           
            MsgBox msg & vbCrLf & vbCrLf & "Created successfully!", , "New Shift"
            'Me.AF3616.SourceObject = "frm3616"
            Dim sid As Integer
            sid = shiftID
            DoCmd.Close acForm, "frmShiftBOARD"
            DoCmd.OpenForm "frmShiftBoard", , , "shiftid = " & sid, acFormEdit, , sid
            Exit Sub
        End If
        
        
'----------------------------------------------------------------------------------------------------

    Case False 'Shift end
        If shiftTitle Like "*swingshift*" Then
            If MsgBox("Close the aerodrome?", vbQuestion + vbYesNo, "Shift Board") = vbYes Then rwy = "AD CLSD"
        End If
        If MsgBox("This will apply your digital signature to all paperwork regarding this shift. Continue?", vbInformation + vbYesNo, "Shift Board") = vbNo Then
            btnStartStop = True
            btnStartStop_AfterUpdate
            Exit Sub
        End If
        
        On Error Resume Next
        If Not UtilForm3616.signLog(shiftID, recsrc, 2) Then
            btnStartStop = True
            btnStartStop_AfterUpdate
            Exit Sub
        End If
        
        Me.closed = True
        If Dirty Then Dirty = False
        DoCmd.Close acForm, "frmShiftBoard"
    End Select
End Sub

Private Sub Detail_Click()
For Each ctl In Controls
    If TypeOf ctl Is ListBox Then
        ctl.Value = ""
    End If
Next
End Sub

Private Sub Form_BeforeUpdate(cancel As Integer)
On Error Resume Next
btnSave = Me.Dirty
    If Not bClose Then
        'If MsgBox("Save?", vbQuestion + vbYesNo, "Shift Board") = vbNo Then
            'cancel = True
            Undo
            Exit Sub
        'End If
    End If
    bClose = False
End Sub

Private Sub Form_Current()
bClose = False
btnStartStop = Not NewRecord
btnSave = Me.Dirty

Dim rs As DAO.Recordset
Dim rsAmoc As DAO.Recordset
Dim offshift() As String
    
    If NewRecord And IsNull(Me.shiftID) Then
        'Get data from previous shift.
        lstOnShift.RowSource = ""
        btnStartStop.Caption = "Start Shift"
        Set rs = CurrentDb.OpenRecordset("SELECT * FROM tblShiftManager WHERE shiftid = " & Nz(DMax("shiftid", "tblshiftmanager"), 0))
        With rs
            On Error Resume Next
            For Each ctl In Me.Controls
                If ctl.Tag Like "*lastShift*" Then
                    If Not IsNull(.Fields(ctl.Name)) Then ctl.Value = .Fields(ctl.Name)
                End If
            Next
        End With
    Else
    
        
        btnStartStop.Caption = "Close and Sign"
        btnSave.Visible = True
        refreshAmocList
        
        AF3616.SourceObject = "frm3616"
    End If
    shiftTitle.Enabled = NewRecord
    tcMain.Visible = Not NewRecord
    
Set rs = Nothing
End Sub

Private Sub Form_Dirty(cancel As Integer)
btnSave = Me.Dirty
bClose = False
End Sub

Private Sub Form_Error(DataErr As Integer, response As Integer)
    If DataErr = 2169 Then response = True
End Sub

Public Function getShiftTitle() As String
getShiftTitle = shiftTitle.Column(0)
End Function

Private Function getSelection(ByRef lst As ListBox) As String()
Dim s() As String
With lst
    For Each i In .ItemsSelected
        ReDim Preserve s(0 To idx) As String
        s(idx) = .Column(0, i)
        idx = idx + 1
    Next
End With

getSelection = s
End Function

Private Sub btnAddAmoc_Click()
If lstAMOCs.ItemsSelected.Count = 0 Then Exit Sub

Dim lstAmocSel() As String
lstAmocSel = getSelection(lstAMOCs)

    lstOnShift.RowSource = IIf(lstOnShift.RowSource = "", "", lstOnShift.RowSource & ";") & join(lstAmocSel, ";")
    amoc = lstOnShift.RowSource
    Dim authLevel As Integer
    For Each a In lstAmocSel
        lstAMOCs.RemoveItem (a)
        authLevel = DLookup("authlevel", "tbluserauth", "opInitials = '" & Right(a, 2) & "'")
        lstOnShiftChange.add (a & " " & IIf(authLevel = 8, "(TRAINEE)", "(AMOC)") & " ON SHIFT")
    Next
    If Left(lstOnShift.RowSource, 5) = "Empty" Then lstOnShift.RemoveItem "Empty"
    
    Me.superLead.SetFocus
    Me.amoc = lstOnShift.RowSource
    Me.Dirty = True
    btnSave = Dirty
    bClose = False
    
    'TODO Handle changes after shift has already started
End Sub

Private Sub Form_Load()
On Error Resume Next
If Nz(shiftID, 0) = 0 Then
    If Nz(Parent.Name) <> "frmShiftDash" And IsNull(Parent.OpenArgs) Then
        'Silent fail
    ElseIf Parent.Name = "frmShiftDash" And Not IsNull(Parent.OpenArgs) Then
        shiftID = Parent.OpenArgs
        'Undo
        bClose = True
        filter = "shiftID = " & shiftID
        FilterOn = True
        btnSave.Visible = True
    End If
End If
End Sub

Private Sub refreshAmocList()
Dim rsAmoc As DAO.Recordset
Dim offshift() As String
'Parse AMOC list
'If Not NewRecord Then

    If Not IsNull(Me.amoc) Then lstOnShift.RowSource = Me.amoc
    
    Set rsAmoc = CurrentDb.OpenRecordset("SELECT tblUserAuth.rank & ' ' & Left([firstname],1) & '. ' & [lastname] & '/' & [opinitials] AS Expr1 FROM tblRank INNER JOIN tblUserAuth ON tblRank.rank = tblUserAuth.rank WHERE (((tblUserAuth.authLevel)<=8)) ORDER BY tblRank.order DESC;")
    With rsAmoc: Do While Not .EOF
    'For Each a In Split(lstOnShift.RowSource, ";")
    If !Expr1 <> Nz(superLead) Then
        ReDim Preserve offshift(0 To idx) As String
        offshift(idx) = !Expr1
        idx = idx + 1
    End If
    'Next
    .MoveNext
    Loop: End With
    
    lstAMOCs.RowSource = join(offshift, ";")
    
    For Each a In Split(lstOnShift.RowSource, ";")
        For Each b In Split(lstAMOCs.RowSource, ";")
            If a = b Then
                lstAMOCs.RemoveItem b
            End If
        Next
    Next
'End If
End Sub

Private Sub Form_Open(cancel As Integer)
For Each ctl In Controls
    If Not TypeOf ctl Is Label Then
        If ctl.Tag Like "*lastshift*" Then
            ctl.OnLostFocus = "=dirtyCheck()"
        End If
    End If
Next
End Sub

Private Sub Form_Timer()
Form.Requery
AF3616.Form.Requery
frmInspections.Form.Requery
btnReloadNOTAMs_Click
End Sub

Private Sub lstAMOCs_Click()
lstOnShift = ""
End Sub

Private Sub lstOnShift_Click()
lstAMOCs = ""
End Sub

Private Sub shiftTitle_AfterUpdate()
Dim rs As DAO.Recordset
If Not IsNull(shiftTitle) Then
    If IsNull(shiftStart) Then
        shiftStart = Date
    End If
    shiftStart = serialDate(shiftStart) & " " & DLookup("shiftStart", "tblShiftDef", "shiftTitle = '" & shiftTitle & "'")
    
    'shiftEnd = Date & " " & DLookup("shiftEnd", "tblShiftDef", "shiftTitle = '" & shiftTitle & "'")
    shiftEnd = DateAdd("n", 60 * DLookup("length", "tblShiftDef", "shiftTitle = '" & shiftTitle & "'"), shiftStart)
    If shiftEnd < shiftStart Then shiftEnd = DateAdd("d", 1, shiftEnd)
    
    tcMain.Visible = True
    tcMain.Pages(1).SetFocus
    
    refreshAmocList

End If
End Sub

Private Sub superLead_AfterUpdate()
Dim authLevel(2) As Integer
If Not NewRecord Then
    authLevel(0) = Nz(DLookup("authlevel", "tbluserauth", "opInitials = '" & Right(Nz(superLead.OldValue), 2) & "'"))
    authLevel(1) = Nz(DLookup("authlevel", "tbluserauth", "opInitials = '" & Right(superLead, 2) & "'"))
    lstOnShiftChange.add superLead.OldValue & " " & IIf(authLevel(0) < 6, "(AMOS)", "(AMSL)") & " OFF SHIFT. " & _
                            superLead & " " & IIf(authLevel(1) < 6, "(AMOS)", "(AMSL)") & " ON SHIFT. "
End If
refreshAmocList
End Sub

Private Sub tcMain_Change()
btnSave = Me.Dirty
bClose = Not Dirty

Select Case tcMain
Case 0:
    AF3616.Form!logView.Requery
End Select
End Sub
