VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmStereoFlightPlan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
Private bClose As Boolean

Private Sub btnDelete_Click()
    If MsgBox("This flight plan will be deleted.", vbInformation + vbOKCancel, "Stereo Manager") = vbOK Then
        log "Stereo plan " & Tag & " deleted.", Me.Name & "btnDelete_click"
        DoCmd.Close acForm, Me.Name
    End If
End Sub

Private Sub btnSave_Click()
    If Dirty Then
        Dirty = False
        log "Stereo plan " & Tag & " updated.", Me.Name & "btnSave_click"
    End If
    
    bClose = True
    DoCmd.Close acForm, Me.Name
End Sub

Private Sub Form_Current()
    bClose = False
End Sub

Private Sub Form_BeforeUpdate(cancel As Integer)
    If Not bClose Then
        'cancel = True
        Undo
        Exit Sub
    End If
End Sub
