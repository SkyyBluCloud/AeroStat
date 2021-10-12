VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'   AeroStat: Airfield Management
'
'
'    Copyright 2017-2022 Shane J. DeSouza
'       https://github.com/SkyyBluCloud/AeroStat
'       sdesouza187@gmail.com
'
'   Licensed under the Apache License, Version 2.0 (the "License");
'   you may not use this file except in compliance with the License.
'   You may obtain a copy of the License at
'
'       http://www.apache.org/licenses/LICENSE-2.0
'
'   Unless required by applicable law or agreed to in writing, software
'   distributed under the License is distributed on an "AS IS" BASIS,
'   WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
'   See the License for the specific language governing permissions and
'   limitations under the License.

Option Compare Database
Dim bClose, forceClose As Boolean
Dim usn, qusn As String

Private Sub switchWindow(ByVal nw As String, Optional ByVal sf As SubForm = Null)
If IsNull(sf) Then Set sf = Me!view
    sf.SourceObject = "loading"
    DoEvents
    sf.SourceObject = nw
End Sub

Private Sub adpm_Click()
    switchWindow "frmADPM", Me!view
End Sub

Private Sub analytics_Click()
    switchWindow "frmTrafficCount", Me!view
End Sub

Private Sub btnMap_Click()
    switchWindow "frmParkingMap", Me!view
End Sub

Private Sub btnRegViewer_Click()
    switchWindow "frmRegViewer", Me!view
End Sub

Private Sub btnReports_Click()
    switchWindow "frmReports", Me!view
End Sub

Private Sub btnSettings_Click()
    switchWindow "frmSettings", Me!view
End Sub

Private Sub feedback_Click()
    DoCmd.OpenForm "frmFeedback", , , , acFormAdd
End Sub

Private Sub log_Click()
On Error Resume Next
Dim isClosed As String
shiftID = DLookup("lastShift", "tblUserAuth", "username = '" & Environ$("username") & "'")
isClosed = Nz(DLookup("closed", "tblShiftManager", "shiftID = " & Nz(shiftID, 0)), True)
    
    If isClosed Then
        DoCmd.OpenForm ("frmShiftSelect")
        Exit Sub
    End If
    
    DoCmd.OpenForm "frm3616", , , , , , shiftID
End Sub

Private Sub menu_Click()
    switchWindow "frmMenu", Me!view
End Sub

Private Sub qrcsop_Click()
    MsgBox "Not ready yet.", vbInformation, "AeroStat"
End Sub

Private Sub Form_Close()
On Error Resume Next
closeAllForms
End Sub

Private Sub lblTitle_Click()
DoCmd.OpenForm "frmAbout"
End Sub

Private Sub notams_Click()
    switchWindow "frmNOTAM", Me!view
End Sub

Private Sub shiftBrief_Click()
'On Error Resume Next
'Dim isClosed As String
Dim shiftID As Integer: shiftID = DLookup("lastShift", "tblUserAuth", "username = '" & Environ$("username") & "'")
Dim isClosed As Boolean: isClosed = Nz(DLookup("closed", "tblShiftManager", "shiftID = " & Nz(shiftID, 0)), True)
Dim onShift As Boolean: onShift = DLookup("onShift", "tbluserauth", "username = '" & Environ$("username") & "'")

    If Not onShift Or isClosed Then
        DoCmd.OpenForm "frmShiftBoard", , , , acFormAdd
        
    ElseIf Not isClosed Then
    
        If CurrentProject.AllForms("frmShiftBoard").IsLoaded Then
        
            With DoCmd
                .SelectObject acForm, "frmShiftBoard"
                .Restore
            End With
            Forms!frmshiftboard.SetFocus
        Else
            DoCmd.OpenForm "frmShiftBoard", , , "shiftid = " & shiftID, acFormEdit, , shiftID
        End If
    End If



'view.SourceObject = "loading"
'
'view.SourceObject = "frmShiftManager"
'view.Form.dashShiftID = IIf(isClosed, 0, shiftID)
End Sub

Private Sub tracker_Click()
    switchWindow "frmTrafficLog", Me!view
End Sub

Private Sub Form_Load()
Dim salute As String
Dim lv As String
Dim RS As DAO.Recordset
'Dim usn As String
'Dim qusn As String



'lv = DLookup("LastVer", "tblSettings")
sp = True
On Error GoTo sperr

'    DoCmd.OpenForm "frmLoading"
'
    
    Set RS = CurrentDb.OpenRecordset(qusn)
    
    With RS
        .edit
        !BOOT = False
        !RS = False
        !spAccess = getAccessSP
        !lastLogin = Now()
        !isLoggedIn = True
        !lastSystem = Environ$("computername")
        
        Forms!frmloading!loadingText.Caption = "Checking WWAs..."
        DoEvents
        
        'DoCmd.OpenForm "WWADaemon", , , , , acHidden, DLookup("wwafeed", "tblsettings")
        
        DoEvents
        
        Forms!frmloading!loadingText.Caption = "Starting AeroStat..."
        DoEvents
        
        view.SourceObject = "frmTrafficLog"
        DoCmd.Close acForm, "frmLoading"
        
'        If Not !lastVersion = lv Then
'            !lastVersion = lv
'            DoCmd.OpenForm "frmChangeLogDisplay", , , "version = '" & lv & "'"
'        End If
            
        .update
        .Close
    End With
    Set RS = Nothing
    
Load_Exit:
    If Date = DateSerial(Year(Date), 7, 12) Then
        With lblEE
            .Caption = "***Happy Birthday XD!!!***"
            .ForeColor = RGB(90, 290, 255)
            .FontBold = True
            .Visible = True
        End With
    End If
    
    DoCmd.RunCommand acCmdAppMaximize
    DoEvents
    
    Exit Sub
    
sperr:
    MsgBox Error$
    Resume Next
End Sub

Public Function getProjName()
    getProjName = Replace(CurrentProject.Name, ".accdb", "")
End Function

Private Sub Form_Open(cancel As Integer)
On Error GoTo open_err
Dim salute As String, lv As Double
    
    'DoCmd.OpenForm "console"
    'Check for first run
    If DLookup("firstrun", "lclver") Then
        'Do first time setup shit
        DoCmd.OpenForm "frmSetup", , , , , acDialog
        If DLookup("firstrun", "lclver") Then
            cancel = True
            Exit Sub
        End If
    End If
    
    If Not CurrentProject.Name Like "*DEV*" Then DoCmd.ShowToolbar "Ribbon", acToolbarNo
    
    DoCmd.OpenForm "frmLoading"
    Forms!frmloading!loadingText.Caption = "Checking environment..."
    DoEvents
    log "Database launched from " & CurrentProject.Path, "frmMain.Form_Open", "WARN"
    
    If InStr(1, CurrentProject.Path, Environ$("username")) = 0 And Not CurrentProject.Name Like "*DEV*" Then
        log "Expected " & Environ$("userprofile"), "frmMain.Form_Open", "WARN"
        Dim dbFrom As String, dbTo As String, fso As New FileSystemObject
        
        'If Left(CurrentProject.Path, "2") = "\\" Then CurrentDb.Execute "UPDATE lclver SET frontEndLocation = " & CurrentProject.fullName
        
        'dbFrom = DLookup("frontEndSource", "tblSettings")
        dbTo = Environ$("userprofile") & "\Documents\AeroStat\"
        'If dir(dbTo, vbNormal) = "" Then createPath dbTo
        Forms!frmloading!loadingText.Caption = "Creating local files..."
        log CurrentProject.fullName, "frmMain.Form_Open"
        DoEvents
        If dir(dbTo, vbNormal) = "" Then fso.CreateFolder dbTo
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
    'If Right(CurrentProject.Name, 5) = "accde" Then CurrentDb.Execute "UPDATE tblSettings SET frontEndSource = '" & CurrentProject.Name & "'"
    
    
    If Util.relinkTables(DLookup("backend", "lclver"), Forms!frmloading) Then
        usn = Nz(DLookup("username", "tblUserAuth", "username = '" & getUSN & "'"))
        qusn = "SELECT * FROM tblUserAuth WHERE username = '" & usn & "'"
        If usn = "" Then
            DoCmd.OpenForm "frmUserInfo", , , , acFormAdd
            cancel = True
            Exit Sub
        End If
    Else
        GoTo open_err
    End If
    
    'DoCmd.Maximize
    
    
    
'    If Not CurrentProject.Name Like "*DEV*" Then DoCmd.ShowToolbar "Ribbon", acToolbarNo
    
'    If InStr(1, CurrentProject.Path, Environ$("userprofile")) = 0 And Not CurrentProject.Name Like "*DEV*" Then
'        Dim dbFrom, dbTo As String
'        Dim fso As New FileSystemObject
'
'        'If Left(CurrentProject.Path, "2") = "\\" Then CurrentDb.Execute "UPDATE lclver SET frontEndLocation = " & CurrentProject.fullName
'
'        dbFrom = DLookup("frontEndSource", "tblSettings")
'        dbTo = Environ$("userprofile") & DLookup("frontEndSource", "tblSettings")
'        If dir(dbTo, vbDirectory) = "" Then createPath dbTo
'        fso.CopyFile dbFrom, dbTo
'        Application.FollowHyperlink dbTo
'        cancel = True
'        Application.Quit
'    Exit Sub
        
        'redo this please....
        'Shell "\\znre-fs-022v\Wing\374OG\374OSS\374OSS.OSAA\00_Section_Info\50_Section_Specific_Info\Traffic Log Tracker\DATABASE\BACKEND\update.bat"
'        MsgBox "This database should be launched from your user profile. Things may not work well.", vbInformation, "AeroStat"
        

'        DBEngine.CompactDatabase dbFrom, dbTo
        'Application.Quit
'    ElseIf DLookup("lclver", "lclver") <> DLookup("lastver", "tblSettings") Then
'        If MsgBox("An update is available. Update now?", vbInformation + vbYesNo, "Update") = vbYes Then
'            Shell "\\znre-fs-022v\Wing\374OG\374OSS\374OSS.OSAA\00_Section_Info\50_Section_Specific_Info\Traffic Log Tracker\DATABASE\BACKEND\update.bat"
'            Application.Quit
'        End If

open_exit:
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
        GoTo open_exit
        
    Case Else
        If MsgBox("An error has occured. The database might not start correctly. Proceed?" & vbCrLf & "(" & err & ") " & Error$, vbCritical + vbYesNo, "AeroStat") = vbNo Then
            ErrHandler err, Error$, Me.Name
            Application.Quit
        End If
End Select
ErrHandler err, Error$, Me.Name
Resume Next
End Sub

Private Sub Form_Timer()
On Error GoTo errtrap
Dim usn As String
Dim qusn As String
Dim RS As DAO.Recordset
usn = Nz(DLookup("username", "tblUserAuth", "username = '" & Environ$("username") & "'"))
qusn = "SELECT * FROM tblUserAuth WHERE username = '" & usn & "'"

'    Set rs = CurrentDb.OpenRecordset(qusn)
'    With rs
'        .edit
'        !lastLogin = Now
'        .update
'        .Close
'    End With
'    Set rs = Nothing
    
    CurrentDb.Execute "UPDATE tblUserAuth SET lastLogin = Now() WHERE username = '" & usn & "'"
        
    If DLookup("BOOT", "tblUserAuth", "username='" & Environ$("username") & "'") Then
        bClose = True
        Application.Quit
        Exit Sub
    ElseIf DLookup("RS", "tbluserauth", "username='" & Environ$("username") & "'") Then
        MsgBox "The database has been updated and will now restart." & vbCrLf & "Applied changes were already saved.", vbOKOnly, "AeroStat"
        Application.FollowHyperlink DLookup("data", "tblSettings", "key = ""dbroot""") & DLookup("data", "tblSettings", "key = ""frontendsource""")
        Application.Quit
        Exit Sub
    End If
sexit:
    Exit Sub
errtrap:
    MsgBox Error$
    ErrHandler err, Error$, Me.Name & ".Form_Timer"
    Resume Next
    
End Sub

Private Sub Form_Unload(cancel As Integer)
On Error GoTo errtrap
Dim usn As String
Dim qusn As String
Dim RS As DAO.Recordset
Dim rsShift As DAO.Recordset
usn = Nz(DLookup("username", "tblUserAuth", "username = '" & Environ$("username") & "'"))
qusn = "SELECT * FROM tblUserAuth WHERE username = '" & usn & "'"
Set RS = CurrentDb.OpenRecordset(qusn)
Set rsShift = CurrentDb.OpenRecordset("SELECT * FROM tblShiftManager WHERE shiftID = " & RS!lastShift)

'        If Not DLookup("closed", "tblshiftmanager", "shiftid = " & rs!lastshift & " AND right(superlead,2) like '" & rs!opInitials & "'") Then
'            If DLookup("authlevel", "tbluserauth", "username = '" & usn & "'") <= 6 Then
'                If MsgBox("The shift is still open. Close shift?", vbQuestion + vbYesNo, "AeroStat") = vbYes Then
'
'
'                    rsShift.edit
'                    rsShift!closed = True
'                    rsShift.update
'
'                End If
'            End If
'        End If
    
        With RS
            .edit
            !isLoggedIn = False
            .update
            .Close
        End With
        Set RS = Nothing
        
sexit:
    Exit Sub
errtrap:
    ErrHandler err, Error$, "frmMain.Form_Unload"
    
End Sub

Public Sub Form_Error(DataErr As Integer, response As Integer)
    If DataErr = 2169 Then
        response = True
    ElseIf err = 3024 Then
        If relinkTables(DLookup("backend", "lclver")) Then MsgBox "Backend loaded.", vbInformation, "AeroStat"
        Exit Sub
    End If
End Sub
