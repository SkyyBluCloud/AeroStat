VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmTrafficCountDetails"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub details_DblClick(cancel As Integer)
    DoCmd.OpenForm "quick_input", , , "ID = " & details.Column(0), acFormEdit
End Sub

Private Sub details_GotFocus()
    details.Requery
End Sub

Private Sub Form_Load()
details.RowSource = OpenArgs
lblTotal.Caption = details.ListCount - 1 & " Flight Plan" & IIf(details.ListCount - 1 = 1, "", "s") & "."
End Sub
