VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmSurfaces"
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
lstSurfaces.SetFocus
End Sub

Private Sub btnDelete_Click()
If MsgBox("Delete?", vbQuestion + vbYesNo, "Surfaces") = vbYes Then
    CurrentDb.Execute "DELETE FROM tblSurfaces WHERE ID = " & lstSurfaces
    lstSurfaces.SetFocus
    lstSurfaces.Requery
End If
End Sub

Private Sub btnNew_Click()
tabCtl = tabCtl + 1
Me.DataEntry = True
End Sub

Private Sub btnSave_Click()
    bClose = True
    If Dirty Then Dirty = False
    'DoCmd.Close acForm, Me.Name
    lstSurfaces.SetFocus
    lstSurfaces.Requery
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

Private Sub lstSurfaces_DblClick(cancel As Integer)
If IsNull(lstSurfaces) Then Exit Sub
tabCtl = 1
Me.filter = "sfcID = " & lstSurfaces
Me.FilterOn = True
End Sub

Private Sub TabCtl_Change()

btnBack.Visible = tabCtl <> 0
End Sub
