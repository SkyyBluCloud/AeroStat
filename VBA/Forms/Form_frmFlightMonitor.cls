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
Dim bClose As Boolean

Private Sub btnDateLeft_Click()
varDate = DateAdd("d", -1, varDate)
lstPPR.Requery
End Sub

Private Sub btnDateRight_Click()
varDate = DateAdd("d", 1, varDate)
lstPPR.Requery
End Sub

Private Sub Form_BeforeUpdate(cancel As Integer)
If Not bClose Then
    Undo
    'Cancel = True
    Exit Sub
End If
End Sub

Private Sub Form_Current()
bClose = False
End Sub

Private Sub Form_Load()
'If Not CurrentProject.Name Like "*DEV*" Then tabCtl.Style = 2
Me.TimerInterval = Nz(getSettings("globalSyncLength"), 5000)

    Dim qdf As DAO.QueryDef: Set qdf = CurrentDb.QueryDefs("qAnnounce")
    With qdf
        .Parameters("cd1") = Date
        .Parameters("cd2") = Date
        Set lstAnnounce.Recordset = .OpenRecordset
    End With
    
    Dim qdf2 As DAO.QueryDef: Set qdf2 = CurrentDb.QueryDefs("qOnStation")
    With qdf2
        .Parameters("bos") = True
        Set lstOnStation.Recordset = .OpenRecordset
    End With
    
    Set qdf = Nothing
End Sub

Private Sub Form_Timer()
    If Not FormSyncUtil.isFormSynced(Me.Name) Then
        Dim ctl: For Each ctl In Me.Controls
            If TypeOf ctl Is ListBox Then
                ctl.Requery
            End If
        Next ctl
        DoEvents
        FormSyncUtil.syncForm Me.Name, True
    End If
End Sub

'Private Sub lstInbound_Click()
'log lstInbound, Me.Name
'End Sub
'
'Private Sub lstOutbound_Click()
'log lstOutbound, Me.Name
'End Sub
'
'Private Sub lstPPR_Click()
'log lstPPR, Me.Name
'End Sub


Private Sub lstPPR_DblClick(cancel As Integer)
DoCmd.OpenForm "new_ppr", , , "PPR = '" & lstPPR & "'", acFormReadOnly, acDialog
End Sub

Private Sub varDate_DblClick(cancel As Integer)
varDate = Date
lstPPR.Requery
End Sub
