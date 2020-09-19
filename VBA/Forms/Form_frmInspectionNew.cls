VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmInspectionNew"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Public bClose As Boolean

Private Function getCheckType(ByVal idx As Integer) As String 'TODO: Should not be hard coded
getCheckType = Array("IFE/GE", "RSC/RCR", "SEVERE WX", "UNAUTHORIZED LANDING", "ADPM VIOLATION", "FOD", "CONSTRUCTION", "BASH", "OTHER")(idx - 1)
End Function

Private Sub add483_Click()
    DoCmd.OpenForm "frm483CompCheck", , , , acFormAdd, acDialog, shiftID
    Remarks.Requery
End Sub

Private Sub addBASH_Click()
DoCmd.OpenForm "frmBASH", , , , , acDialog, Name
End Sub

Private Sub ctlInTime_Exit(cancel As Integer)
If Not IsNull(ctlInTime) Then
    If Len(ctlInTime) <> 4 Then
        cancel = True
        ctlInTime.BorderColor = RGB(255, 0, 0)
    Else
        ctlInTime.BorderColor = RGB(166, 166, 166)
    End If
End If
End Sub

Private Sub ctlOutTime_Exit(cancel As Integer)
If Not IsNull(ctlOutTime) Then
    If Len(ctlOutTime) <> 4 Then
        cancel = True
        ctlOutTime.BorderColor = RGB(255, 0, 0)
    Else
        ctlOutTime.BorderColor = RGB(166, 166, 166)
    End If
End If
End Sub

Private Sub done_Click()
On Error GoTo errtrap
Dim looped As Boolean
Dim firstEntryComplete As Boolean
Dim rsLog As DAO.Recordset
Dim entry As String
Dim ctSimple As String
Dim t As Date
    
'    If IsNull(outTime) Then
'        outTime.BorderColor = RGB(255, 0, 0)
'        If IsNull(inTime) Then
'            If Not NewRecord Then
'                inTime.BorderColor = RGB(255, 0, 0)
'                Exit Sub
'            End If
'        End If
'        Exit Sub
'    End If

    For Each ctl In Controls
        If TypeOf ctl Is CommandButton Or CheckBox Then ctl.Enabled = False
    Next
    
    Dim fail As Boolean
    For Each ctl In Controls
        If TypeOf ctl Is TextBox Then
            Select Case ctl.Name
                Case "callsign", "initials"
                    If IsNull(ctl.Value) Then
                        ctl.BorderColor = RGB(255, 0, 0)
                        fail = True
                    Else
                        ctl.BorderColor = RGB(166, 166, 166)
                    End If
                    ctl.Value = UCase(ctl.Value)
                
                Case "ctlouttime"
                    If IsNull(ctl.Value) Or Len(ctl.Value) > 4 Then
                        ctl.BorderColor = RGB(255, 0, 0)
                        fail = True
                    Else
                        ctl.BorderColor = RGB(166, 166, 166)
                    End If
                    
                Case "ctlintime"
                    If (IsNull(ctl.Value) Or Len(ctl.Value) > 4) _
                    And Not NewRecord Then
                        ctl.BorderColor = RGB(255, 0, 0)
                        fail = True
                    Else
                        ctl.BorderColor = RGB(166, 166, 166)
                    End If
            End Select
        End If
    Next
    If Nz(checkType) = "" And specialCheck = 0 Then
        ctBox.BorderColor = RGB(255, 0, 0)
        fail = True
    Else
        ctBox.BorderColor = RGB(166, 166, 166)
    End If
    If fail Then GoTo fb
    
    
    outTime = checkDate & " " & getTime4Char(Replace(ctlOutTime, ":", ""))
    If Not IsNull(ctlInTime) Then inTime = checkDate & " " & getTime4Char(Replace(ctlInTime, ":", ""))
    
    'TODO: Should not be hard coded......
    If specialCheck <> 0 Then
        Select Case specialCheck
            Case 1: ctSimple = "OPENING CK"
            Case 2: ctSimple = "DLY INSP"
            Case 3: ctSimple = "LGTG CK"
        End Select
    Else
        For Each c In Split(checkType, ",")
            If c = 9 Then
                If Nz(Remarks) = "" Then
                    Remarks.BorderColor = RGB(255, 0, 0)
                    GoTo fb
                Else
                    Remarks.BorderColor = RGB(166, 166, 166)
                End If
            End If
            ctSimple = ctSimple & IIf(Nz(ctSimple) = "", "", "/") & IIf(c = 9, Remarks, getCheckType(c))
        Next
    End If
    

log:
    entry = Callsign & "/" & Initials & " ON AFLD " & IIf(specialCheck <> 0, "FOR " & ctSimple, "CK (" & ctSimple & ")")
    'outTime = Format(Left(entryTime, 2) & "." & Right(entryTime, 2), "hh:nn") WTF???
    t = checkDate & " " & ZToL(outTime)
    If firstEntryComplete Or (Not NewRecord And logOff = 0) Then
        firstEntryComplete = True
        If logOff = 0 Then
            entry = Callsign & "/" & Initials & " OFF AFLD. " & IIf(Nz(Remarks) = Nz(Remarks.OldValue), "NO CHANGES TO REPORT", Nz(Remarks))
        Else
            entry = Callsign & "/" & Initials & " OFF AFLD. " & Nz(Remarks, "NO CHANGES TO REPORT")
        End If
        'entry = Callsign & "/" & Initials & " OFF AFLD. " & IIf(Nz(Remarks) = "", "NO CHANGES TO REPORT", Nz(Remarks))
        'entry = Callsign & "/" & Initials & " OFF AFLD. " & IIf(Nz(Remarks) = "", "NO CHANGES TO REPORT", "")
        t = checkDate & " " & ZToL(inTime)
        firstEntry = False
        closed = True
    End If
    
    If bAddLog Then
        Set rsLog = CurrentDb.OpenRecordset("SELECT * FROM tbl3616 WHERE shiftID = " & shiftID & " AND ID = " & IIf(Not firstEntryComplete, logOn, logOff))
        With rsLog
            If .EOF Then
                .AddNew
            Else
                .edit
            End If
            !shiftID = shiftID
            !entryTime = LToZ(t)
            !entry = entry
            !Initials = DLookup("opinitials", "tbluserauth", "username = '" & Environ$("username") & "'")
            .update
            .Bookmark = .LastModified
            IIf(Not firstEntryComplete, logOn, logOff) = !ID
        End With
    Else
        logOn = -1
    End If
    
    If Not IsNull(inTime) And Not firstEntryComplete Then
        firstEntryComplete = True
        bClose = True
        If Dirty Then Dirty = False
        GoTo log
    End If
    
    bClose = True
    If Dirty Then Dirty = False
    
sexit:
    'Parent.Parent.Form!sf3616.Form.logView.Requery
    
    Parent!back.Visible = False
    'DoCmd.save acForm, Me.Name
    Parent.changeView "frmInspectionList", shiftID
    'DoCmd.close acForm, "frmInspections"
    Exit Sub
fb:
    For Each ctl In Controls
        If TypeOf ctl Is CommandButton Or CheckBox Then ctl.Enabled = True
    Next
    Exit Sub
    Resume Next
errtrap:
    If err = 3101 Then
        'MsgBox "This 483 is not in the database.", vbInformation, "AF483"
    End If
    ErrHandler err, Error$, Me.Name
    GoTo fb
End Sub

Private Sub Form_Load()
On Error Resume Next
'    If IsNull(outTime) Then
'        outTime = LToZ(Format(Now, "hh:nn"))
'    ElseIf IsNull(inTime) Then
'        inTime = LToZ(Format(Now, "hh:nn"))
'    End If
    Me.shiftID = Parent.shiftID
    Parent.back.Tag = "frmInspectionList"
    Parent.back.Visible = True
    
End Sub

Private Sub Form_Open(cancel As Integer)
    For Each ctl In Controls
        If ctl.Tag = "ct" Then
            ctl.OnClick = "=updateCheckType()"
        End If
    Next
End Sub

Private Function loadCheckType()
    If Nz(checkType) <> "" Then
        For Each ctl In Controls
            If TypeOf ctl Is CheckBox Then
                For Each t In Split(checkType, ",")
                    If t = Right(ctl.Name, 1) And ctl.Tag = "ct" Then
                        ctl.Value = True
                        Exit For
                    End If
                Next
            End If
        Next
    ElseIf Nz(Remarks) = "" And specialCheck <> 0 Then
        Select Case specialCheck
            Case 1: Remarks = "RSC ###, BWC ###, BARRIERS ###"
            Case 2: Remarks = "DAILY INSPECTION CMPLT"
            Case 3: Remarks = "LIGHTING CK CMPLT"
        End Select
    End If
End Function

Private Function updateCheckType()
    checkType = ""
    specialCheck = 0
    For Each ctl In Controls
        If TypeOf ctl Is CheckBox And Left(ctl.Name, 2) = "op" Then
            If ctl.Value = True And ctl.Tag = "ct" Then
                checkType = checkType & IIf(IsNull(checkType), "", ",") & Right(ctl.Name, 1)
            End If
        End If
    Next
        
End Function

Private Sub Inspection_AfterUpdate()
    checklist.Visible = Inspection
End Sub

Private Sub Form_BeforeUpdate(cancel As Integer)
    If Not bClose Then
        cancel = True
        Undo
        Exit Sub
    End If
End Sub

Private Sub Form_Current()
bClose = False

'    If IsNull(outTime) Then
'        ctlOutTime = Format(LToZ(Now), "hhnn")
'    ElseIf IsNull(ctlInTime) Then
'        ctlInTime = Format(LToZ(Now), "hhnn")
'    End If
    ctlOutTime = Format(Nz(outTime, LToZ(Now)), "hhnn")
    If Not NewRecord Then
        ctlInTime = Format(Nz(inTime, LToZ(Now)), "hhnn")
        loadCheckType
    End If

    
    add483.Visible = Not NewRecord
    addBASH.Visible = Not NewRecord
    
End Sub

Private Sub Form_Error(DataErr As Integer, response As Integer)
    If DataErr = 2169 Or DataErr = 3101 Then response = True
End Sub

Private Sub specialCheck_Click()
Dim scName As String

    If DCount("ID", "tblinspections", "specialCheck = " & specialCheck & " AND shiftID = " & shiftID) > 0 Then 'Special check was already logged
        lblError.Caption = "Warning: An inspection/special check was already logged for this shift."
        lblError.Visible = True
    Else
        lblError.Visible = False
    End If
    
    For Each ctl In Controls
        If TypeOf ctl Is CheckBox And ctl.Tag = "ct" Then
            ctl.Value = False
        End If
    Next
End Sub
