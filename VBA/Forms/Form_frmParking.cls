VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmParking"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
Dim bClose As Boolean

Public Function getSfcType(ByVal i As Integer) As Variant
'1;RWY;2;TWY;3;APRON;4;Service Road;5;Misc (ACFT);6;Misc (Other)

    Select Case i
        Case 1: getSfcType = "RWY"
        Case 2: getSfcType = "TWY"
        Case 3: getSfcType = "APRON"
        Case 4: getSfcType = "Service Road"
        Case 5: getSfcType = "Misc (ACFT)"
        Case 6: getSfcType = "Misc (Other)"
        Case Else: getSfcType = Null
    End Select
End Function

Private Sub btnCancel_Click()
DoCmd.Close acForm, Me.Name
End Sub

'Private Sub btnAddEquipment_Click()
'Dim newEquip As String: newEquip = InputBox("Enter the name of the equipment you would like to add:", Nz(sfcName, "Airfield") & " Equipment", "Ex. BAK12 or DASR")
'CurrentDb.Execute "INSERT INTO '" & Me.equipment.RowSource & "' ("
'equipment.Requery
'End Sub

Private Sub btnBack_Click()
tabCtl = 0
End Sub

Private Sub btnNew_Click()
tabCtl = tabCtl + 1
Me.DataEntry = True
End Sub

Private Sub btnSave_Click()
    bClose = True
    If Dirty Then Dirty = False
    'DoCmd.Close acForm, Me.Name
End Sub

Private Sub btnDeleteRamp_Click()
    If NewRecord Then
        Undo
        tabCtl = 0
    ElseIf MsgBox("This will also delete the associated parking spots. Continue?", vbQuestion + vbYesNo, "Delete Ramp") = vbYes Then
        tabCtl = 0
        CurrentDb.Execute "DELETE * FROM tblParkingRamp INNER JOIN tblParkingSpot on tblParkingSpot.rampID = tblParkingRamp.rampID WHERE tblParkingRamp.rampID = " & lstRampID
        lstRampID.Requery
    End If
End Sub

Private Sub btnNewRamp_Click()
Me.tabCtl = 1 'Ramp
Me.DataEntry = True
End Sub

Private Sub btnNewSpot_Click()
Dim newSpot As String: newSpot = InputBox("Enter a name for the new Spot:", "New Spot", "Ex: 'A1'")
If IsNull(lstRampID) Then Exit Sub
    
    Do While Nz(newSpot) <> ""
        CurrentDb.Execute "INSERT INTO qParkingSpot (tblParkingSpot.rampID, tblParkingSpot.spot) VALUES (" & lstRampID & ", '" & newSpot & "')", dbFailOnError
        lstSpot.Requery
        newSpot = InputBox("Enter a name for the new Spot:", "New Spot", "Ex: 'A1'")
    Loop
    
sexit:
    Exit Sub
errtrap:
    ErrHandler err, Error$, Me.Name & ".btnNewSpot_Click"
End Sub

Private Sub btnNewRow_Click()
tabCtl = 2 'New Row
Me.DataEntry = True
End Sub

Private Sub btnRowSave_Click()
If Dirty Then Dirty = False
tabCtl = 0
lstSpot.Requery
End Sub

Private Sub btnSaveRamp_Click()
On Error GoTo errtrap
Dim fail As Boolean

    Dim ctl: For Each ctl In Controls
        If ctl.Tag = "required" Then
            If Nz(ctl.Value) = "" Then
                ctl.BorderColor = vbRed
                fail = True
            Else
                ctl.BorderColor = RGB(166, 166, 166)
            End If
        End If
    Next ctl
    If fail Then Exit Sub
    
    bClose = True
    If Dirty Then Dirty = False
    tabCtl = 0
    lstRampID.Requery
sexit:
    Exit Sub
errtrap:
    ErrHandler err, Error$, Me.Name & ".btnSaveRamp_Click"
    
End Sub

Private Sub Closed_AfterUpdate()
Select Case Me.closed
Case True
    Me.closed.Caption = "CLOSED"
Case False
    Me.closed.Caption = "OPEN"
End Select
End Sub

Private Sub Form_BeforeUpdate(cancel As Integer)
If Not bClose Then
    'cancel = True
    Undo
    Exit Sub
End If
End Sub

Private Sub Form_Current()
bClose = False
End Sub

Private Sub Form_Error(DataErr As Integer, response As Integer)
    If DataErr = 2169 Then response = True
End Sub

Private Sub Form_Load()
If Not CurrentProject.Name Like "*DEV*" Then tabCtl.Style = 2 'TabStyle "None"
End Sub

Private Sub lstRampID_AfterUpdate()
    lstSpot.Requery
End Sub

Private Sub TabCtl_Change()
btnBack.Visible = tabCtl <> 0
End Sub
