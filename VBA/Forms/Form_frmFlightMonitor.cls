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
Dim fpReloadTime As Integer

Private Sub btnDateLeft_Click()
varDate = DateAdd("d", -1, varDate)
lstPPR.Requery
End Sub

Private Sub btnDateRight_Click()
varDate = DateAdd("d", 1, varDate)
lstPPR.Requery
End Sub

Private Sub Detail_Click()
Dim c: For Each c In Me.Controls
    If TypeOf c Is ListBox Then c.Value = ""
Next c
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
Me.TimerInterval = Nz(getSettings("globalSyncLength"), 10000)
    
    CurrentDb.Execute "UPDATE tblUserAuth SET BOOT = false, RS = false, spAccess = " & getAccessSP & ", lastlogin = now(), " & _
                        "isLoggedIn = True, lastSystem = '" & Util.getWorkstation & "' WHERE username = '" & Util.getUser & "'"
        
    Forms!frmloading!loadingText.Caption = "Checking WWAs..."
    DoEvents
    
    'DoCmd.OpenForm "WWADaemon", , , , , acHidden, DLookup("wwafeed", "tblsettings")
    
    DoEvents
    
    Forms!frmloading!loadingText.Caption = "Starting AeroStat..."
    DoEvents

    Dim qdf As DAO.QueryDef: Set qdf = CurrentDb.QueryDefs("qAnnounce")
    With qdf
        .Parameters("cd1") = Date
        .Parameters("cd2") = DateAdd("d", 5, Date)
        Set lstAnnounce.Recordset = .OpenRecordset
    End With
    
    Dim qdf2 As DAO.QueryDef: Set qdf2 = CurrentDb.QueryDefs("qOnStation")
    With qdf2
        .Parameters("bos") = True
        Set lstOnStation.Recordset = .OpenRecordset
    End With
    
    Dim qdf3 As DAO.QueryDef: Set qdf3 = CurrentDb.QueryDefs("qFlightOrder")
    With qdf3
        '.Parameters("varFlow") = "*"
        Set lstInbound.Recordset = .OpenRecordset
    End With
        
    Dim c: For Each c In Me.Controls
        If TypeOf c Is ListBox Then c.Value = ""
    Next c
    
    
    
    Set qdf = Nothing
    Set qdf2 = Nothing
    DoCmd.Close acForm, "frmloading"
    DoCmd.RunCommand acCmdAppMaximize
    DoEvents
End Sub

Private Sub Form_Open(cancel As Integer)
    If Not CurrentProject.Name Like "*DEV*" Then DoCmd.ShowToolbar "Ribbon", acToolbarNo
    
    DoCmd.OpenForm "frmLoading"
    Forms!frmloading!loadingText.Caption = "Checking environment..."
    
    log "Database launched from " & CurrentProject.Path, "frmMain.Form_Open", "WARN"
    DoEvents
    
    'If InStr(1, CurrentProject.Path, Util.getUser) = 0 And Not CurrentProject.Name Like "*DEV*" Then
    If InStr(1, CurrentProject.Path, Util.getUser) = 0 Then
        log "Expected " & Environ$("userprofile"), "frmMain.Form_Open", "WARN"
        Dim dbFrom As String, dbTo As String, fso As New FileSystemObject
        
        If Right(CurrentProject.Name, 5) = "accde" Then _
            CurrentDb.Execute "UPDATE tblSettings SET data = '" & CurrentProject.fullName & "' WHERE key = 'frontEndSource'"
        
        'dbFrom = DLookup("frontEndSource", "tblSettings")
        dbTo = Environ$("userprofile") & "\Documents\AeroStat\"
        'If dir(dbTo, vbNormal) = "" Then createPath dbTo
        Forms!frmloading!loadingText.Caption = "Creating local files..."
        log CurrentProject.fullName, "frmMain.Form_Open"
        DoEvents
        
        If dir(dbTo, vbDirectory) = "" Then fso.CreateFolder dbTo
        fso.CopyFile CurrentProject.fullName, dbTo & CurrentProject.Name
        Forms!frmloading!loadingText.Caption = "Activating..."
        DoEvents
        
        Application.FollowHyperlink dbTo & CurrentProject.Name
        cancel = True
        Application.Quit
        Exit Sub
    End If
    
    Forms!frmloading!loadingText.Caption = "Loading global settings..."
    DoEvents
    
    If Util.relinkTables(DLookup("backend", "lclver"), Forms!frmloading) Then
        Dim qusn: qusn = "SELECT * FROM tblUserAuth WHERE username = '" & Util.getUser & "'"
        If CurrentDb.OpenRecordset(qusn).EOF Then
            DoCmd.OpenForm "frmUserInfo", , , , acFormAdd
            cancel = True
            Exit Sub
        End If
    Else
        GoTo open_err
    End If
sexit:
    Exit Sub
    
open_err:


Select Case err
    Case 3043
        MsgBox "Unable to connect to database (ShareDrive not detected). Please try again later, or restart your computer." & vbCrLf & "(3043)", vbCritical, "AeroStat"
        'Application.Quit
        
    Case 3024, 3044
        If Not relinkTables(DLookup("backend", "lclver")) Then
            cancel = True
            Exit Sub
        End If
        
        Form_Open 0
        Exit Sub
        
    Case 7867
        GoTo sexit
        
    Case Else
        If MsgBox("An error has occured. The database might not start correctly. Proceed?" & vbCrLf & "(" & err & ") " & Error$, vbCritical + vbYesNo, "AeroStat") = vbNo Then
            ErrHandler err, Error$, Me.Name
            Application.Quit
        End If
End Select
ErrHandler err, Error$, Me.Name
    
End Sub

Private Sub Form_Timer()
    If Not FormSyncUtil.isFormSynced(Me.Name) Or fpReloadTime >= DLookup("data", "tblsettings", "key = 'globalSyncInterval'") Then
    
        lblUpdating.Visible = True
        DoEvents
        
        If Not FormSyncUtil.isFormSynced(Me.Name) Then Util.sndPlaySound getSettings("soundUpdate"), SND_ASYNC
        
        Dim ctl: For Each ctl In Me.Controls
            If TypeOf ctl Is ListBox Then
                ctl.Requery
            End If
        Next ctl
        DoEvents
        
        FormSyncUtil.syncForm Me.Name, True
        fpReloadTime = -1
    End If
    fpReloadTime = fpReloadTime + 1
    
    lblUpdating.Visible = False
    DoEvents
    
    CurrentDb.Execute "UPDATE tblUserAuth SET isLoggedIn = True, lastLogin = Now() WHERE username = '" & Util.getUser & "'"
End Sub

Private Sub Form_Unload(cancel As Integer)
CurrentDb.Execute "UPDATE tblUserAuth SET isLoggedIn = False where username = '" & Util.getUser & "'"
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

Private Sub tabCtl_Click()
Dim c: For Each c In Me.Controls
    If TypeOf c Is ListBox Then c.Value = ""
Next c
End Sub

Private Sub tglHideLCL_AfterUpdate()
Dim qdf As DAO.QueryDef: Set qdf = CurrentDb.QueryDefs("qFlightOrder")

    With qdf
        lblUpdating.Visible = True
        DoEvents
        
        Select Case tglHideLCL
        Case True
            Set lstInbound.Recordset = CurrentDb.OpenRecordset("SELECT * FROM qFlightOrder WHERE Flow <> 'LCL'")
            tglHideLCL.Caption = "Show LCL"
        Case False
            
            Set lstInbound.Recordset = CurrentDb.OpenRecordset("qFlightOrder")
            tglHideLCL.Caption = "Hide LCL"
        End Select
    End With

    lblUpdating.Visible = False
    DoEvents
End Sub


Private Sub varDate_DblClick(cancel As Integer)
varDate = Date
lstPPR.Requery
End Sub
