VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmRegViewerEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Dim bClose As Boolean

Private Sub btnDelete_Click()
    CurrentDb.Execute "DELETE FROM [@PUBS] WHERE ID = " & lstRegs
    lstRegs.Requery
End Sub

Private Sub btnNew_Click()
    FilterOn = False
    DataEntry = True
End Sub

Private Sub btnSave_Click()
    bClose = True
    If Dirty Then Dirty = False
    lstRegs.Requery
    lblSaved.Visible = True
    Me.TimerInterval = 2000
End Sub

Private Sub Form_BeforeUpdate(cancel As Integer)
    If Not bClose Then Undo
End Sub

Private Sub Form_Current()
    bClose = False
End Sub

Private Sub Form_Timer()
    lblSaved.Visible = False
    Me.TimerInterval = 0
End Sub

Private Sub lstRegs_AfterUpdate()
    filter = "ID = " & lstRegs
    FilterOn = True
End Sub


