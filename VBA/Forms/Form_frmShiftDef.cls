VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmShiftDef"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub btnAddNew_Click()
    Me.DataEntry = True
End Sub

Private Sub btnDelete_Click()
    CurrentDb.Execute "DELETE FROM tblShiftDef WHERE ID = " & lstShifts
    Me.FilterOn = False
    lstShifts.Requery
End Sub

Private Sub btnSave_Click()
    If Dirty Then Dirty = False
    lstShifts.Requery
    lstShifts = ID
End Sub

Private Sub lstShifts_Click()
    Me.filter = "ID = " & lstShifts
    Me.FilterOn = True
End Sub
