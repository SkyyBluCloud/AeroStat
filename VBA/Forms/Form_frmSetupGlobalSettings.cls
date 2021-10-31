VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmSetupGlobalSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
Dim bClose As Boolean

Private Sub btnDone_Click()
bClose = True
If Dirty Then Dirty = False
End Sub

Private Sub Form_BeforeUpdate(cancel As Integer)
bClose = False
End Sub

Private Sub Form_Current()
    If Not bClose Then
        Undo
    End If
End Sub

Private Sub lstSettings_AfterUpdate()
Me.Filter = "ID = " & lstSettings
Me.FilterOn = True
End Sub
