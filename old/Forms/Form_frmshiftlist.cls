VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmshiftlist"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub shiftView_DblClick(cancel As Integer)
If IsNull(shiftView.Column(0)) Then Exit Sub
Parent.changeView "frmShiftDetails", shiftView.Column(0)
End Sub
