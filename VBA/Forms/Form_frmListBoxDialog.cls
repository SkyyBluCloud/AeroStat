VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmListBoxDialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub ctlListBox_DblClick(cancel As Integer)
Me.Visible = False
End Sub

Private Sub Form_Load()
ctlListBox.RowSource = OpenArgs
End Sub

Private Sub Form_Open(cancel As Integer)
cancel = IsNull(OpenArgs)
End Sub
