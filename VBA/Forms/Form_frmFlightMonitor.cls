VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmFlightMonitor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit



Private Sub Form_Load()
If Not CurrentProject.Name Like "*DEV*" Then tabCtl.Style = 2
End Sub

Private Sub Form_Timer()
    If Not FormSyncUtil.isFormSynced(Me.Name) Then
        Dim ctl: For Each ctl In Me.Controls
            If TypeOf ctl Is ListBox Then
                ctl.Requery
            End If
        Next ctl
        
        FormSyncUtil.syncForm Me.Name, True
    End If
End Sub

Private Sub varDate_AfterUpdate()
    lstInbound.Requery
End Sub

