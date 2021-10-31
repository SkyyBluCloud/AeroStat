VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmShiftNew"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
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
    bClose = False
    If Not NewRecord Then
        If closed Then
            lblClosed.Visible = True
            FormHeader.BackColor = RGB(166, 166, 166)
            FormFooter.BackColor = RGB(166, 166, 166)
        Else
            closeShift.Visible = True
        End If
    End If
End Sub

Private Sub btnLog_Click()
    If Not NewRecord Then
        If DLookup("closed", "tblShiftManager", "shiftID = " & shiftID) Then
            If Not DLookup("opInitials", "tblUserAuth", "username = '" & Util.getUser & "'") = Right(DLookup("superLead", "tblShiftManager", "shiftID = " & shiftID), 2) Then
                MsgBox "This shift was closed. Only the AMOS/AMSL can edit the log.", vbInformation, "AeroStat"
                'Exit Sub
            End If
        End If
        DoCmd.OpenForm "frm3616", , , , , , shiftID
    End If
End Sub

Private Sub closeShift_Click()
If NewRecord Then Exit Sub
Dim authLevel As Integer
Dim usn As String
usn = Util.getUser
authLevel = DLookup("authlevel", "tbluserauth", "username = '" & usn & "'")

    If authLevel > 6 Then
        MsgBox "Only the AMOS can close the shift.", vbInformation, "Shift Manager"
        Exit Sub
    End If
    
    If Not IsNull(amosSig) Then
        MsgBox "This shift was already signed.", vbInformation, "Shift Manager"
        Exit Sub
    End If
    
    If Not IsNull(namoSig) Or Not IsNull(afmSig) Then
        MsgBox "This shift was already finalized by the NAMO or AFM.", vbInformation, "Shift Manager"
        Exit Sub
    End If

    If MsgBox("This will lock all paperwork related to the shift, and apply your digital signature to the reports. After this, only yourself, the NAMO, and the AFM can make changes. Continue?", vbQuestion + vbYesNo, "Shift Manager") = vbYes Then
        closed = True
        amosSig = usn
        amosSigTime = Now
        lblClosed.Visible = True
        FormHeader.BackColor = RGB(166, 166, 166)
        FormFooter.BackColor = RGB(166, 166, 166)
        done_Click
    End If
    
End Sub

Private Sub done_Click()
Dim RS As DAO.Recordset
Dim inc As Boolean
bClose = True
    
    If Nz(shiftTitle) = "" Then
        shiftTitle.BorderColor = RGB(255, 0, 0)
        Exit Sub
    Else
        shiftTitle.BorderColor = RGB(166, 166, 166)
    End If
    
    es = DLookup("shiftID", "tblShiftManager", "Format(shiftStart,'dd-mmm-yy hh:nn') = '" & shiftStart & "'")
    If Not IsNull(es) Then
        If MsgBox("This shift was already started. Would you like to view it?", vbQuestion + vbYesNo, "AeroStat") = vbYes Then
            Undo
            DoCmd.OpenForm "frmShiftNew", , , "shiftID = " & es, acFormEdit
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
            If InStr(1, ctl.Tag, "nonull") > 0 Then
                If Nz(ctl.Value) = "" Then
                    inc = True
                    ctl.BorderColor = RGB(255, 0, 0)
                Else
                    ctl.BorderColor = RGB(166, 166, 166)
                End If
            End If
            If InStr(1, ctl.Tag, "caps") > 0 Then
                ctl.Value = UCase(ctl.Value)
            End If
        End If
    Next
    If inc Then GoTo fb
    
    Set RS = CurrentDb.OpenRecordset("tblUserAuth")
    Dim joined As String
    Dim arrJoined() As String
    Dim msg As String
    With RS
    
        For Each ctl In Controls
            If TypeOf ctl Is ComboBox And InStr(1, ctl.Tag, "init") > 0 Then
                Do While Not .EOF
                    If !opInitials = Right(ctl.Value, 2) Then
                        .edit
                        !lastShift = Me.shiftID
                        joined = joined & ctl.Value & ";"
                        .update
                        Exit Do
                    Else
                        .MoveNext
                    End If
                Loop
            End If
            .MoveFirst
        Next
    .Close
    End With
    
    If NewRecord Then
        arrJoined = Split(joined, ";")
        msg = "Shift: " & shiftTitle & " (" & shiftStart & " - " & shiftEnd & ")" & vbCrLf & _
                "AMOS/AMSL: " & arrJoined(0) & vbCrLf & _
                "AMOC: "
        
        For i = 1 To UBound(arrJoined)
            msg = msg & " " & arrJoined(i)
            If Not i = UBound(arrJoined) Then msg = msg & ","
        Next

        Dirty = False
    
        '''''AUTOLOG
        Set RS = CurrentDb.OpenRecordset("SELECT * FROM tbl3616 WHERE tbl3616.shiftID = " & Me.shiftID & ";")
        If RS.RecordCount = 0 Then
        al1 = Nz(DLookup("authlevel", "tbluserauth", "opInitials = '" & Right(superLead, 2) & "'"))
        al2 = Nz(DLookup("authlevel", "tbluserauth", "opInitials = '" & Right(Nz(amoc1), 2) & "'"))
        al3 = Nz(DLookup("authlevel", "tbluserauth", "opInitials = '" & Right(Nz(amoc2), 2) & "'"))
        al4 = Nz(DLookup("authlevel", "tbluserauth", "opInitials = '" & Right(Nz(amoc4), 2) & "'"))
        With RS
            .AddNew
            !shiftID = Me.shiftID
            !entryTime = LToZ(Me.shiftStart)
            !Initials = DLookup("opInitials", "tblUserAuth", "username = '" & Util.getUser & "'")
            'amoc field should store full name instead of initials
            '!entry = IIf(shiftTitle = "DAYSHIFT", "LOG OPEN. ", "") ...
            !entry = "LOG OPEN. " & UCase(superLead) & " " & getPos(DLookup("authlevel", "tbluserauth", "opInitials = '" & Right(superLead, 2) & "'")) & _
            IIf(Not IsNull(amoc1), IIf(Not IsNull(amoc2), ", ", " AND ") & UCase(Nz(amoc1)) & " " & getPos(IIf(al2 = 5 Or al2 = 6, 7, al2)), "") & _
            IIf(Not IsNull(amoc2), IIf(Not IsNull(amoc3), ", ", " AND ") & UCase(Nz(amoc2)) & " " & getPos(IIf(al3 = 5 Or al3 = 6, 7, al3)), "") & _
            IIf(Not IsNull(amoc3), IIf(Not IsNull(amoc4), ", ", " AND ") & UCase(Nz(amoc3)) & " " & getPos(IIf(al4 = 5 Or al4 = 6, 7, al4)), "") & _
            IIf(Not IsNull(amoc4), " AND " & UCase(Nz(amoc4)) & " " & getPos(Nz(DLookup("authlevel", "tbluserauth", "opInitials = '" & Right(Nz(amoc4), 2) & "'"))), "") & " ON DUTY. " & _
            "NOTAMS: " & notamM & "M, " & notamV & "V, " & notamL & "L. FLT PLANNING ROOM CK CMPLT. " & _
            IIf(shiftTitle = "SWINGSHIFT", "RWY " & rwy & " IN USE. RSC " & rsc & ", BWC " & bwc & ". BARRIERS " & IIf(barrier = "WITHDRAWN", "WITHDRAWN", "CONFIGURED FOR RWY " & barrier) & ". ", "") & _
            shiftTitle & " CHECKLIST INITIATED. MONITORING RAMPNET FOR PROPER PHRASEOLOGY THROUGHOUT SHIFT"
            'FUTURE: User customization
            .update
        End With
        End If
       
        MsgBox msg & vbCrLf & vbCrLf & "Created successfully!", , "New Shift"
        Forms!frmMain.view.SourceObject = "frmShiftManager"
        'DoCmd.close acForm, "frmShiftSelect"
    End If
    
sexit:
    tor = Now
    'DoCmd.close acForm, "frmShiftNew"
    Exit Sub
fb:
'    For Each ctl In Controls
'        If TypeOf ctl Is TextBox And InStr(1, ctl.Tag, "init") > 0 Then ctl.Value = Right(ctl.Value, 2)
'    Next
End Sub

Private Sub Form_Load()
nshiftID = Parent.Parent.getShiftID
    If nshiftID <> 0 Then
        Undo
        Filter = "shiftID = " & nshiftID
        FilterOn = True
    End If
    'If NewRecord Then log "New Record."
End Sub

Private Sub newBASH_Click()
    If Not NewRecord Then DoCmd.OpenForm "frmBASH", , , , acFormAdd, , shiftID
End Sub

Private Sub recount_Click()
notamM = DCount("NOTAM", "qNOTAMM", "left([NOTAM],1) = 'M'")
notamL = DCount("NOTAM", "qNOTAML", "left([NOTAM],1) = 'L'")
done_Click
End Sub

Private Sub shiftTitle_AfterUpdate()

    
    If IsNull(shiftStart) Then
        shiftStart = Date
    End If
    shiftStart = serialDate(shiftStart) & " " & DLookup("shiftStart", "tblShiftDef", "shiftTitle = '" & shiftTitle & "'")
    
    'shiftEnd = Date & " " & DLookup("shiftEnd", "tblShiftDef", "shiftTitle = '" & shiftTitle & "'")
    shiftEnd = DateAdd("n", 60 * DLookup("length", "tblShiftDef", "shiftTitle = '" & shiftTitle & "'"), shiftStart)
    If shiftEnd < shiftStart Then shiftEnd = DateAdd("d", 1, shiftEnd)

'    Select Case shiftTitle
'        Case "DAYSHIFT"
'            shiftStart = Date & " 5.00"
'            shiftEnd = Date & " 13.15"
'        Case "SWINGSHIFT"
'            shiftStart = Date & " 13.15"
'            shiftEnd = Date & " 22.00"
'        Case "MIDSHIFT"
'            shiftStart = Date & " 22.00"
'            shiftEnd = DateAdd("h", 7, shiftStart)
'    End Select
End Sub
