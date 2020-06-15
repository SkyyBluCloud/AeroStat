VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'   AeroStat: Airfield Management(TM)
'
'
'    Copyright 2017-2020 Shane J. DeSouza(SkyyBluCloud)
'       skyyblucloud@gmail.com
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

Private Sub btnReports_Click()
    switchWindow "frmBackOffice", Me!view
End Sub

Private Sub feedback_Click()
    DoCmd.OpenForm "frmFeedback", , , , acFormAdd
End Sub

Private Sub log_Click()
On Error Resume Next
Dim isclosed As String
shiftID = DLookup("lastShift", "tblUserAuth", "username = '" & Environ$("username") & "'")
isclosed = Nz(DLookup("closed", "tblShiftManager", "shiftID = " & Nz(shiftID, 0)), True)
    
    If isclosed Then
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

Private Sub lblTitle_Click()
DoCmd.OpenForm "frmAbout", Me!view
End Sub

Private Sub notams_Click()
    switchWindow "frmNOTAM", Me!view
End Sub

Private Sub shiftBrief_Click()
'On Error Resume Next
'Dim isClosed As String
Dim shiftID As Integer: shiftID = DLookup("lastShift", "tblUserAuth", "username = '" & Environ$("username") & "'")
Dim isclosed As Boolean: isclosed = Nz(DLookup("closed", "tblShiftManager", "shiftID = " & Nz(shiftID, 0)), True)
Dim onShift As Boolean: onShift = DLookup("onShift", "tbluserauth", "username = '" & Environ$("username") & "'")

    If Not onShift Or isclosed Then
        DoCmd.OpenForm "frmShiftSelect"
        Exit Sub
    ElseIf Not isclosed Then
        DoCmd.OpenForm "frmShiftBoard", , , "shiftid = " & shiftID, acFormEdit, , shiftID
    End If



'view.SourceObject = "loading"
'
'view.SourceObject = "frmShiftManager"
'view.Form.dashShiftID = IIf(isClosed, 0, shiftID)
End Sub

Private Sub tracker_Click()
    switchWindow "Traffic Log", Me!view
End Sub

Private Sub Form_Load()
Dim salute As String
Dim lv As String
Dim rs As DAO.Recordset
'Dim usn As String
'Dim qusn As String


'lv = DLookup("LastVer", "tblSettings")
sp = True
On Error GoTo sperr

'    DoCmd.OpenForm "frmLoading"
'
    
    Set rs = CurrentDb.OpenRecordset(qusn)
    
    With rs
        .edit
        !BOOT = False
        !rs = False
        !spAccess = getAccessSP
        !lastLogin = Now()
        !isLoggedIn = True
        !lastSystem = Environ$("computername")
        
        DoCmd.RunCommand acCmdAppMaximize
        Forms!frmLoading!loadingText.Caption = "Starting AeroStat..."
        
        view.SourceObject = "Traffic Log"
        DoCmd.Close acForm, "frmLoading", acSaveYes
        
'        If Not !lastVersion = lv Then
'            !lastVersion = lv
'            DoCmd.OpenForm "frmChangeLogDisplay", , , "version = '" & lv & "'"
'        End If
            
        .Update
        .Close
    End With
    Set rs = Nothing
    
Load_Exit:
    Exit Sub
    
sperr:
    MsgBox Error$
    Resume Next
End Sub

Private Sub Form_Open(cancel As Integer)
On Error GoTo open_err
Dim salute As String, lv As Double
    
    'Check for first run
    If DLookup("firstrun", "lclver") Then
        'Do first time setup shit
        DoCmd.OpenForm "frmSetup", , , , , acDialog
        If DLookup("firstrun", "lclver") Then
            cancel = True
            Exit Sub
        End If
    End If
    
    If Not CurrentProject.Name Like "*MASTER*" Then DoCmd.ShowToolbar "Ribbon", acToolbarNo
    
    If InStr(1, CurrentProject.Path, Environ$("userprofile")) = 0 And Not CurrentProject.Name Like "*MASTER*" Then
        Dim dbFrom, dbTo As String, fso As New FileSystemObject
        
        'If Left(CurrentProject.Path, "2") = "\\" Then CurrentDb.Execute "UPDATE lclver SET frontEndLocation = " & CurrentProject.fullName
        
        dbFrom = DLookup("frontEndSource", "tblSettings")
        dbTo = Environ$("userprofile") & "\Documents\AeroStat\"
        'If dir(dbTo, vbNormal) = "" Then createPath dbTo
        If dir(dbTo, vbNormal) = "" Then fso.CreateFolder dbTo
        fso.CopyFile dbFrom, dbTo
        Application.FollowHyperlink dbTo & CurrentProject.Name
        cancel = True
        Application.Quit
        Exit Sub
    End If
    
    DoCmd.OpenForm "frmLoading"
    Forms!frmLoading!loadingText.Caption = "Loading global settings..."
    
    If Util.relinkTables(DLookup("backend", "lclver"), Forms!frmLoading) Then
        usn = Nz(DLookup("username", "tblUserAuth", "username = '" & Environ$("username") & "'"))
        qusn = "SELECT * FROM tblUserAuth WHERE username = '" & usn & "'"
        If usn = "" Then
            DoCmd.OpenForm "frmUserInfo", , , , acFormAdd
            cancel = True
            Exit Sub
        End If
    Else
        GoTo open_err
    End If
    
    
    
    
    
'    If Not CurrentProject.Name Like "*MASTER*" Then DoCmd.ShowToolbar "Ribbon", acToolbarNo
    
'    If InStr(1, CurrentProject.Path, Environ$("userprofile")) = 0 And Not CurrentProject.Name Like "*MASTER*" Then
'        Dim dbFrom, dbTo As String
'        Dim fso As New FileSystemObject
'
'        'If Left(CurrentProject.Path, "2") = "\\" Then CurrentDb.Execute "UPDATE lclver SET frontEndLocation = " & CurrentProject.fullName
'
'        dbFrom = DLookup("frontEndSource", "tblSettings")
'        dbTo = Environ$("userprofile") & DLookup("frontEndRuntime", "tblSettings")
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
Dim rs As DAO.Recordset
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
        bClose = True
'        Shell "\\znre-fs-022v\Wing\374OG\374OSS\374OSS.OSAA\00_Section_Info\50_Section_Specific_Info\Traffic Log Tracker\DATABASE\BACKEND\update.bat"
'        Application.Quit

        Dim dbFrom, dbTo As String
        Dim fso As New FileSystemObject
        
        'If Left(CurrentProject.Path, "2") = "\\" Then CurrentDb.Execute "UPDATE lclver SET frontEndLocation = " & CurrentProject.fullName
        
        dbFrom = DLookup("frontEndSource", "tblSettings")
        dbTo = Environ$("userprofile") & DLookup("frontEndRuntime", "tblSettings")
        'If dir(dbTo, vbNormal) = "" Then createPath dbTo
        If dir(dbTo, vbNormal) = "" Then fso.CreateFolder dbTo
        fso.CopyFile dbFrom, dbTo
        Application.FollowHyperlink dbTo & CurrentProject.Name
        Application.Quit
        Exit Sub
    End If
sExit:
    Exit Sub
errtrap:
    MsgBox Error$
    ErrHandler err, Error$, Me.Name & ".Form_Timer"
    Resume Next
    
End Sub

Private Sub Form_Unload(cancel As Integer)
'On Error Resume Next
Dim usn As String
Dim qusn As String
Dim rs As DAO.Recordset
Dim rsShift As DAO.Recordset
usn = Nz(DLookup("username", "tblUserAuth", "username = '" & Environ$("username") & "'"))
qusn = "SELECT * FROM tblUserAuth WHERE username = '" & usn & "'"
Set rs = CurrentDb.OpenRecordset(qusn)
Set rsShift = CurrentDb.OpenRecordset("SELECT * FROM tblShiftManager WHERE shiftID = " & rs!lastShift)

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
    
        With rs
            .edit
            !isLoggedIn = False
            .Update
            .Close
        End With
        Set rs = Nothing
    
End Sub

Public Sub Form_Error(DataErr As Integer, response As Integer)
    If DataErr = 2169 Then
        response = True
    ElseIf err = 3024 Then
        If relinkTables(DLookup("backend", "lclver")) Then MsgBox "Backend loaded.", vbInformation, "AeroStat"
        Exit Sub
    End If
End Sub
