VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frm483CompCheck"
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

Private Sub compliant_Click()
reason.Visible = Not compliant
    Select Case compliant
        Case True
            compliant.Caption = "Compliant"
            reason = ""
        Case False
            compliant.Caption = "NOT Compliant"
    End Select
End Sub

Private Sub done_Click()
On Error GoTo errtrap
Dim resp As String
bClose = True
Dim rsLog As DAO.Recordset

    Dim fail As Boolean
    For Each ctl In Controls: With ctl
        If TypeOf ctl Is TextBox And .Tag Like "*required*" Then
            If Nz(.Value) = "" Then
                fail = True
                .BorderColor = RGB(255, 0, 0)
            Else
                .BorderColor = RGB(166, 166, 166)
            End If
        End If
    End With: Next
    If fail Then Exit Sub
    
    If IsNull(DLookup("af483", "tbl483", "af483 = '" & AF483 & "'")) Then
        If MsgBox("Confirm information:" & vbCrLf & vbCrLf & _
            AF483 & vbCrLf & _
            rank & " " & lastName & ", " & firstName & " " & Unit, vbInformation + vbYesNo, "AF483") = vbYes Then
            
            Dim appendData() As String
            Dim idx As Integer
            Dim copyDetails As String
            Dim copyCompliant As Boolean
            Dim copyReason As String
            Dim copyShift As Integer
            
            copyDetails = Me.details
            copyCompliant = Me.compliant
            copyReason = Nz(Me.reason)
            copyShift = Me.shiftID
            
            Do Until idx = 6
                For Each ctl In Controls
                    If ctl.Tag Like "*appendData*" And Right(ctl.Tag, 1) = idx Then
                        ReDim Preserve appendData(0 To idx) As String
                        appendData(idx) = ctl.Value
                        idx = idx + 1
                    End If
                Next
            Loop
            'CurrentDb.Execute "INSERT INTO tbl483 (af483,rankID,lastname,firstname,midinit,unit) " & _
                                "SELECT '" & cAF483 & "', '" & rankID & "', '" & lastName & "', '" & firstName & "', '" & midInit & "', '" & Unit & "'", dbFailOnError
            CurrentDb.Execute "INSERT INTO tbl483 (af483,rankID,lastname,firstname,midinit,unit,issuedBy,issueDate,lastCheck) " & _
                                "SELECT '" & join(appendData, "', '") & "', '" & Util.getOpInitials & "', now(), now()", dbFailOnError

            Undo

            Me.AF483 = appendData(0)
            Me.details = copyDetails
            Me.compliant = copyCompliant
            Me.reason = copyReason
            Me.shiftID = copyShift
            Me.by = Util.getOpInitials
            If Dirty Then Dirty = False
        End If
    End If

'DoCmd.RunCommand acCmdSaveRecord
    
    If CurrentProject.AllForms("frmInspections").IsLoaded Then 'Was triggered from inspection form. That will handle the log entry; send log details
        With Forms!frmInspections!view.Form
            !AF483Count = !AF483Count + 1
            resp = !AF483Count & " 483" & IIf(!AF483Count > 1, "s", "") & " LOGGED " & IIf(compliant, "IN COMPLIANCE", "NOT IN COMPLIANCE")
            If InStr(1, !Remarks, !AF483Count - 1 & " 483" & IIf(!AF483Count - 1 > 1, "s", "") & " LOGGED " & IIf(compliant, "IN COMPLIANCE", "NOT IN COMPLIANCE")) > 0 Then
                !Remarks = Replace(!Remarks, !AF483Count - 1 & " 483" & IIf(!AF483Count - 1 > 1, "s", "") & " LOGGED " & IIf(compliant, "IN COMPLIANCE", "NOT IN COMPLIANCE"), resp)
            Else
                !Remarks = Nz(!Remarks) & IIf(IsNull(!Remarks), resp, ". " & resp)
            End If
        End With
        'rmk = Forms!frmInspections!view.Form!Remarks & " 483 CK LOGGED " & IIf(compliant, "IN COMPLIANCE", "NOT IN COMPLIANCE")
        
    ElseIf CurrentProject.AllForms("frm3616").IsLoaded Then 'Was triggered directly from the log (483 check at the counter?)
        If IsNull(DLookup("af483", "tbl483", "af483 = '" & AF483 & "'")) Then
            If MsgBox("Add log entry?", vbQuestion + vbYesNo, "483 Check") = vbYes Then
                Set rsLog = CurrentDb.OpenRecordset("SELECT * FROM tbl3616 WHERE shiftid = " & shiftID)
                With rsLog
                    .AddNew
                    !shiftID = shiftID
                    !entryTime = LToZ(checkDate)
                    !entry = "483 CK LOGGED " & IIf(compliant, "IN COMPLIANCE", "NOT IN COMPLIANCE")
                    !Initials = DLookup("opinitials", "tbluserauth", "username = '" & Environ$("username") & "'")
                    .Update
                End With
            End If
        End If
    End If
    
sExit:
    DoCmd.Close acForm, "frm483CompCheck", acSaveYes
    Exit Sub
errtrap:
'    If err = 3101 Then
'        If MsgBox("Confirm information:" & vbCrLf & vbCrLf & _
'        AF483 & vbCrLf & _
'        rankID & " " & lastName & ", " & firstName & " " & Unit, vbInformation + vbYesNo, "AF483") = vbYes Then
'
'            Dim appendData() As String
'            Dim idx As Integer
'            Do Until idx = 6
'                For Each ctl In Controls
'                    If ctl.Tag Like "*appendData*" And Right(ctl.Tag, 1) = idx Then
'                        ReDim Preserve appendData(0 To idx) As String
'                        appendData(idx) = ctl.Value
'                        idx = idx + 1
'                    End If
'                Next
'            Loop
'
'            CurrentDb.Execute "INSERT INTO tbl483 (af483,rankID,lastname,firstname,midinit,unit) " & _
'                                "SELECT '" & join(appendData, "', '") & "'"
'            If Dirty Then Dirty = False
'        End If
        
'    End If
    
    ErrHandler err, Error$, Me.Name
    Resume Next
End Sub

Private Sub Form_BeforeUpdate(cancel As Integer)
    If Not bClose Then
        cancel = True
        Undo
        Exit Sub
    End If
End Sub

Private Sub Form_Current()
    bClose = False
End Sub

Private Sub Form_Error(DataErr As Integer, response As Integer)
    If DataErr = 2169 Or DataErr = 3101 Then response = True
End Sub

Private Sub Form_Load()
    shiftID = OpenArgs
    by = DLookup("opinitials", "tbluserauth", "username = '" & Environ$("username") & "'")
End Sub

Private Sub Form_Open(cancel As Integer)
    If IsNull(OpenArgs) Then
        MsgBox "This form was opened illegally!" & vbCrLf & "OpenArgs: " & Nz(OpenArgs, "NULL!"), vbCritical, "AeroStat"
        'Cancel = True
    End If
End Sub
