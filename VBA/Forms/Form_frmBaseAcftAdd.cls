VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmBaseAcftAdd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub btnUpdate_Click()
    DoCmd.Close acForm, "frmBaseAcftAdd"
End Sub

Private Sub Form_Load()
    Dim s() As String
    s = Split(OpenArgs, ";")
    acType = s(0)
    Tail = s(1)
End Sub
