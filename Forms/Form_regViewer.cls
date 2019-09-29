VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_regViewer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub Form_Load()
    regViewer.ControlSource = OpenArgs
End Sub

Private Sub Form_Open(cancel As Integer)
    If Nz(OpenArgs) = "" Then cancel = True
End Sub
