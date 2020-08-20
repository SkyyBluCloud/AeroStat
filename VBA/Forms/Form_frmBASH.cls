VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmBASH"
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

Private Sub Form_Current()
bClose = False
End Sub

Private Sub Form_Load()
    If NewRecord Then
        shiftID = OpenArgs
        obTime = Now
    End If
End Sub

Private Sub save_Click()
On Error GoTo errtrap
bClose = True
Dim RS As DAO.Recordset
Dim e As Boolean
Dim msg As String
    For Each ctl In Controls
        If TypeOf ctl Is TextBox Then
            If Nz(ctl.Value) = "" And ctl.Name <> "remarks" Then
                ctl.BorderColor = RGB(255, 0, 0)
                e = True
            Else
                ctl.BorderColor = RGB(166, 166, 166)
            End If
        End If
    Next
    If e Then Exit Sub
        
    msg = "BASH: " & Number
    Select Case species
        Case 1
            msg = msg & " Sparrow"
        Case 2
            msg = msg & " Egret"
        Case 3
            msg = msg & " Mynah Bird"
        Case 4
            msg = msg & " Plower"
        Case 5
            msg = msg & " Sea Gull"
        Case 6
            msg = msg & " Rice Bird"
        Case 7
            msg = msg & " Dove"
        Case 8
            msg = msg & " Crow"
        Case 9
            msg = msg & " Black Bird"
        Case 10
            msg = msg & " Hawk"
        Case 11
            ib = InputBox("Please enter the name of the BASH found:", "BASH Report")
            msg = msg & " " & ib
            Remarks = IIf(Not Nz(Remarks) = "", Remarks & vbCrLf, Remarks) & "TYPE/" & ib
    End Select
    msg = msg & IIf(Number > 1, "s", "") & " FOUND"
    
    Select Case activity
        Case 1
            msg = msg & " Flying"
        Case 2
            msg = msg & " Soaring"
        Case 3
            msg = msg & " Loafing"
        Case 4
            msg = msg & " Feeding"
        Case 5
            msg = msg & " Nesting"
    End Select
    msg = msg & " AT " & location & ". Dispersed using"
    
    Select Case dispersalMethod
        Case 1
            msg = msg & " Vehicle"
        Case 2
            msg = msg & " Sirens"
        Case 3
            msg = msg & " Pistol"
        Case 4
            msg = msg & " BASH cannons"
        Case 5
            msg = msg & " Shotgun"
    End Select
    If Not Nz(Remarks) = "" Then
        msg = msg & ". " & Remarks
    End If
    msg = UCase(Trim(msg))
    
    If OpenArgs = "frmInspectionNew" Then
        With Forms!frmInspectionNew
            !Remarks = !Remarks & " " & msg
        End With
        
    Else
        Set RS = CurrentDb.OpenRecordset("SELECT * FROM tbl3616 WHERE shiftID = " & shiftID)
        With RS
        If MsgBox("Add log entry to current shift?", vbQuestion + vbYesNo, "BASH Report") = vbYes Then
            If Not DLookup("closed", "tblShiftManager", "shiftid = " & shiftID) Then
                log "Adding 3616 entry", Me.Name & ".onClick"
                .AddNew
                !shiftID = shiftID
                !entryTime = LToZ(obTime)
                !entry = msg
                !Initials = Initials
                .update
                .Close
            Else
                MsgBox "Either you are not currently in a shift, or the log for the shift has been closed.", vbInformation, "BASH Report"
            End If
        End If
        End With
        Set RS = Nothing
    End If
    'DoCmd.RunCommand acCmdSaveRecord
    DoCmd.Close acForm, "frmBASH", acSaveYes
    MsgBox "BASH entry successfully saved.", vbInformation, "BASH Report"
    
sexit:
    Exit Sub
errtrap:
    ErrHandler err, Error$, Me.Name & ".save_Click"
    
End Sub
