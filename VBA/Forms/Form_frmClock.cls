VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmClock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub Form_Timer()
    For Each ctl In Me.Controls
        If TypeOf ctl Is TextBox Then
            With ctl
                '.enabled = True
                .Requery
                DoEvents
                '.enabled = False
            End With
        End If
    Next ctl
    DoEvents
End Sub
