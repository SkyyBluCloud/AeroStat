VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmShiftDash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub Form_Load()
If Not IsNull(Me.OpenArgs) Then
    Me.view.Form!shiftID = OpenArgs
    Me.view.Form.bClose = True
    With Me.view.Form
        .Undo
        .filter = "shiftid = " & Me.OpenArgs
        .FilterOn = True
    End With
End If
End Sub
