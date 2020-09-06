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

Private Sub btnCancel_Click()
DoCmd.Close acForm, Me.Name
End Sub

Private Sub btnSave_Click()
    bClose = True
    If Dirty Then Dirty = False
    DoCmd.Close acForm, Me.Name
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
End Sub

Private Sub Form_Error(DataErr As Integer, response As Integer)
    If DataErr = 2169 Then response = True
End Sub
