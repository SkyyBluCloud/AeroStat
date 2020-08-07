VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmEnforceAcType"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Public response As String

Private Sub btnCancel_Click()
    If MsgBox("Not choosing a type from this list might cause the traffic count to be wrong. Continue?", vbQuestion + vbYesNo, "A/C Type") = vbNo Then Exit Sub
    response = OpenArgs
    Me.Visible = False
End Sub

Private Sub lstAcType_DblClick(cancel As Integer)
    response = Nz(lstAcType)
    Me.Visible = False
End Sub

