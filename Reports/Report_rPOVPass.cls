VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Report_rPOVPass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub print_Click()
On Error Resume Next
    DoCmd.RunCommand acCmdPrint
End Sub

Private Sub print2_Click()
print_Click
End Sub
