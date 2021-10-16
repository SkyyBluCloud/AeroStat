VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frm3616Template"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Dim bClose As Boolean

Private Sub btnDelete_Click()
On Error Resume Next
    DoCmd.RunCommand acCmdDeleteRecord
    lstTemplates.Requery
End Sub

Private Sub btnNew_Click()
Me.DataEntry = True
End Sub

Private Sub btnSave_Click()
bClose = True
If Dirty Then Dirty = False
lstTemplates.Requery
End Sub

Private Sub Form_BeforeUpdate(cancel As Integer)
If Not bClose Then
    Undo
    Exit Sub
End If
End Sub

Private Sub Form_Current()
bClose = False
End Sub

Private Sub Form_Open(cancel As Integer)
Me.DataEntry = True
End Sub

Private Sub lstTemplates_Click()
Me.filter = "ID = " & lstTemplates
Me.FilterOn = True
End Sub
