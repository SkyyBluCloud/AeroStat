VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmAMC"
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
End Sub

Private Sub Command11_Click()
    Dim fail(2) As Boolean
    If Tail Like "N/A" Then Tail = ""
    For Each ctl In Controls
        If TypeOf ctl Is TextBox Then
            If ctl.Name = "Callsign" And (Not isCallsign Or Nz(ctl.Value) = "") Then
                ctl.BorderColor = RGB(255, 0, 0)
                fail(1) = True
            ElseIf ctl.Name = "Tail" And Nz(ctl.Value) = "" Then
                ctl.BorderColor = RGB(255, 0, 0)
                fail(2) = True
            Else
                ctl.BorderColor = RGB(166, 166, 166)
            End If
        End If
    Next
    If fail(1) And fail(2) Then
        MsgBox "A valid Callsign or Tail Number is required to issue this PPR." & vbCrLf & "(If you must issue this PPR without this information, enter the Tail as 'TBD')", vbCritical, "PPR"
        Exit Sub
    End If
    
    For Each ctl In Controls
        If TypeOf ctl Is TextBox Then
            ctl.Value = Replace(Nz(ctl.Value), vbCrLf, "")
            ctl.Value = Trim(ctl.Value)
        End If
    Next

    With Forms!new_ppr
        If Nz(!Remarks) = "" Then !Remarks = "MSN #: " & MSN & vbCrLf
        !Callsign = IIf(isCallsign, Callsign, Right(MSN, 7))
        !arrDate = ZToL(parseDateToJ(Nz(adate), True))
        !depDate = ZToL(parseDateToJ(Nz(ddate), True))
        !Type = Me.Type
        !depPoint = depPoint
        !Destination = Destination
        !MSN = MSN
        !Tail = Tail
        If IsNull(!Fuel) Then !Fuel = "TBD"
        If IsNull(!Spot) Then !Spot = "AMC"
        For Each ctl In .Controls
            If TypeOf ctl Is TextBox Then ctl.Requery
        Next
        DoCmd.Close acForm, "frmAMC"
    End With
sexit:
    Exit Sub
fb:
    Exit Sub
End Sub

Private Function parseDateToJ(ByVal loc As String, Optional ByVal reverse As Boolean) As Variant
    If Nz(loc) = "" Then Exit Function
    If Len(loc) = 8 Then loc = Left(loc, 4) & "/" & Right(loc, 4)
    If Not reverse Then
        parseDateToJ = Right(Year(LToZ(loc)), 1) & Format(DateDiff("d", "1/1/" & Year(LToZ(loc)), LToZ(loc)) + 1, "000") & "/" & Format(LToZ(loc), "hhnn")
    Else
        parseDateToJ = Format(DateSerial(Left(Year(Now), 3) & Left(loc, 1), 1, 0) + Mid(loc, 2, 3) & " " & Mid(loc, 6, 2) & ":" & Right(loc, 2), "dd-mmm-yy hh:nn")
    End If
End Function

Private Sub Form_Load()
    With Forms!new_ppr
        adate = parseDateToJ(Nz(!arrDate))
        ddate = parseDateToJ(Nz(!depDate))
        Callsign = !Callsign
        isCallsign = Nz(Callsign) <> "" Or Not Nz(Right(!MSN, 7)) = Nz(!Callsign)
        Callsign.enabled = isCallsign
        Me.Type = !Type
        depPoint = !depPoint
        Destination = !Destination
        MSN = !MSN
        Tail = !Tail
    End With
End Sub

Private Sub isCallsign_AfterUpdate()
    Callsign.enabled = isCallsign
    If isCallsign Then Callsign = Forms!new_ppr!Callsign
End Sub
