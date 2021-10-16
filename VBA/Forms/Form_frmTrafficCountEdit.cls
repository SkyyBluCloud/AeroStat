VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmTrafficCountEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub btnSave_Click()
    If Me.Dirty Then
        Me.Dirty = False
        Me.TimerInterval = 1000
        lblSave.Visible = True
    End If
End Sub

Private Sub Form_Timer()
    lblSave.Visible = False
    Me.TimerInterval = 0
End Sub

Private Sub lstItems_AfterUpdate()
    Me.filter = "ID = " & lstItems
    Me.FilterOn = True
End Sub
