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
Dim bClose As Boolean
Dim forceClose As Boolean

Private Sub adpm_Click()
    Me!view.SourceObject = "loading"
    DoEvents
    Me!view.SourceObject = "frmADPM"
End Sub

Private Sub analytics_Click()
    Me!view.SourceObject = "loading"
    DoEvents
    Me!view.SourceObject = "frmTrafficCount"
End Sub

Private Sub btnMap_Click()
    Me!view.SourceObject = "loading"
    DoEvents
    Me!view.SourceObject = "frmParkingMap"
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
    Me!view.SourceObject = "loading"
    DoEvents
    Me!view.SourceObject = "frmMenu"
End Sub

Private Sub qrcsop_Click()
    MsgBox "Not ready yet.", vbInformation, "AeroStat"
End Sub

Private Sub lblTitle_Click()
DoCmd.OpenForm "frmAbout"
End Sub

Private Sub notams_Click()
    Me!view.SourceObject = "loading"
    DoEvents
    Me!view.SourceObject = "frmNOTAM"
End Sub

Private Sub shiftBrief_Click()
'On Error Resume Next
'Dim isClosed As String
Dim shiftID As Integer: shiftID = DLookup("lastShift", "tblUserAuth", "username = '" & Environ$("username") & "'")
isClosed = Nz(DLookup("closed", "tblShiftManager", "shiftID = " & Nz(shiftID, 0)), True)
Dim onShift As Boolean: onShift = DLookup("onShift", "tbluserauth", "username = '" & Environ$("username") & "'")

    If Not onShift Or isClosed Then
        DoCmd.OpenForm "frmShiftSelect"
        Exit Sub
    ElseIf Not isClosed Then
        DoCmd.OpenForm "frmShiftBoard", , , "shiftid = " & shiftID, acFormEdit, , shiftID
    End If



'view.SourceObject = "loading"
'DoEvents
'view.SourceObject = "frmShiftManager"
'view.Form.dashShiftID = IIf(isClosed, 0, shiftID)
End Sub

Private Sub tracker_Click()
    Me!view.SourceObject = "loading"
    DoEvents
    Me!view.SourceObject = "Traffic Log"
End Sub

Private Sub Form_Load()
Dim salute As String
Dim lv As String
Dim rs As DAO.Recordset
Dim usn As String
Dim qusn As String
usn = Nz(DLookup("username", "tblUserAuth", "username = '" & Environ$("username") & "'"))
qusn = "SELECT * FROM tblUserAuth WHERE username = '" & usn & "'"

lv = DLookup("LastVer", "settings")
sp = True
On Error GoTo sperr

    DoCmd.OpenForm "frmLoading"
    DoEvents
    
    Set rs = CurrentDb.OpenRecordset(qusn)
    With rs
        .edit
        !BOOT = False
        !rs = False
        !spAccess = getAccessSP
        !lastLogin = Now()
        !isLoggedIn = True
        !lastSystem = Environ$("computername")
        
        DoCmd.Close acForm, "frmLoading", acSaveYes
        
        If Not !lastVersion = lv Then
            !lastVersion = lv
            DoCmd.OpenForm "frmChangeLogDisplay", , , "version = '" & lv & "'"
        End If
            
        .update
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
Dim salute As String
Dim lv As Double
Dim usn As String
Dim qusn As String
usn = Nz(DLookup("username", "tblUserAuth", "username = '" & Environ$("username") & "'"))
qusn = "SELECT * FROM tblUserAuth WHERE username = '" & usn & "'"
    
    If usn <> "1516576541A" Then DoCmd.ShowToolbar "Ribbon", acToolbarNo
    
    If InStr(1, CurrentProject.Path, Environ$("userprofile")) = 0 Then
        'redo this please....
        Shell "\\znre-fs-022v\Wing\374OG\374OSS\374OSS.OSAA\00_Section_Info\50_Section_Specific_Info\Traffic Log Tracker\DATABASE\BACKEND\update.bat"
        MsgBox "This database should be launched from your user profile. Things may not work well.", vbInformation, "AeroStat"
        'Application.Quit
'    ElseIf DLookup("lclver", "lclver") <> DLookup("lastver", "settings") Then
'        If MsgBox("An update is available. Update now?", vbInformation + vbYesNo, "Update") = vbYes Then
'            Shell "\\znre-fs-022v\Wing\374OG\374OSS\374OSS.OSAA\00_Section_Info\50_Section_Specific_Info\Traffic Log Tracker\DATABASE\BACKEND\update.bat"
'            Application.Quit
'        End If
    End If

    
    If usn = "" Then
        DoCmd.OpenForm "frmUserInfo", , , , acFormAdd
        cancel = True
        Exit Sub
    End If
    
open_exit:
    Exit Sub
    
open_err:
    If err = 3043 Then
        MsgBox "Unable to connect to database (ShareDrive not detected). Please try again later, or restart your computer." & vbCrLf & "(3043)", vbCritical, "AeroStat"
        'Application.Quit
    ElseIf err = 3024 Or err = 3044 Then
        If Not relinkTables Then
            cancel = True
            Exit Sub
        End If
        Form_Open 0
        Exit Sub
    Else
        errHandler err, Error$, Me.Name
    End If
    Resume Next
    
End Sub

Private Sub Form_Timer()
On Error GoTo errTrap
Dim usn As String
Dim qusn As String
Dim rs As DAO.Recordset
usn = Nz(DLookup("username", "tblUserAuth", "username = '" & Environ$("username") & "'"))
qusn = "SELECT * FROM tblUserAuth WHERE username = '" & usn & "'"

    Set rs = CurrentDb.OpenRecordset(qusn)
    With rs
        .edit
        !lastLogin = Now
        .update
        .Close
    End With
    Set rs = Nothing
        
    If DLookup("BOOT", "tblUserAuth", "username='" & Environ$("username") & "'") Then
        bClose = True
        Application.Quit
        Exit Sub
    ElseIf DLookup("RS", "tbluserauth", "username='" & Environ$("username") & "'") Then
        bClose = True
        Shell "\\znre-fs-022v\Wing\374OG\374OSS\374OSS.OSAA\00_Section_Info\50_Section_Specific_Info\Traffic Log Tracker\DATABASE\BACKEND\update.bat"
        Application.Quit
        Exit Sub
    End If
sexit:
    Exit Sub
errTrap:
    MsgBox Error$
    errHandler err, Error$, Me.Name & ".Form_Timer"
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
            .update
            .Close
        End With
        Set rs = Nothing
    
End Sub

Public Sub Form_Error(DataErr As Integer, Response As Integer)
    If DataErr = 2169 Then
        Response = True
    ElseIf err = 3024 Then
        If relinkTables Then MsgBox "Backend loaded.", vbInformation, "AeroStat"
        Exit Sub
    End If
End Sub

Private Sub paperwork_Click()
    MsgBox "Not ready yet.", vbExclamation, "AeroStat"
End Sub
