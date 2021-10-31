VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Report_new3616"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
Public args As Variant
Dim pageCount() As Integer

Private Sub printClick(ByVal d As Variant)
If IsNull(d) Or Not IsDate(d) Then Exit Sub
DoCmd.OpenReport "new3616", acViewReport, , , , d
'DoEvents
'DoCmd.SelectObject acReport, "new3616"
'
'DoCmd.RunCommand acCmdPrint
'DoEvents
'DoCmd.Close acReport, "new3616"
End Sub



Private Sub btnPrint_Click()
On Error GoTo errtrap
    If Not CurrentProject.AllReports("new3616").IsLoaded Then
        DoCmd.OpenReport "new3616", acViewReport, , , acHidden, rDate
        DoEvents
    End If
    DoCmd.SelectObject acReport, "new3616"
    DoCmd.RunCommand acCmdPrint
    DoEvents
    
sexit:
    DoCmd.Close acReport, "new3616"
    Exit Sub
errtrap:
    Select Case err
        Case Else
            ErrHandler err, Error$, Me.Name & ".btnPrint_Click"
            GoTo sexit
    End Select
End Sub

Private Sub Detail_Print(cancel As Integer, PrintCount As Integer)
Dim eol As Double
eol = r3616Shift.Height

Me.Line (Me.Line94.Left, 0)-(Me.Line94.Left, eol)
Me.Line (Me.Line93.Left, 0)-(Me.Line93.Left, eol)
Me.Line (Me.Line92.Left, 0)-(Me.Line92.Left, eol)
Me.Line (Me.sCount.Left - 1, eol - 1)-(r3616Shift.Left, eol - 1)

End Sub

Private Sub namoBox_Click()
Dim authLevel As Integer

authLevel = DLookup("authLevel", "tbluserauth", "username = '" & Util.getUser & "'")
If authLevel > 4 Then Exit Sub

    'if msgbox("By signing this document, you certify that all entries are correct and the form contains all required entries as

End Sub

Private Sub Report_Load()
On Error Resume Next
Dim rsShift As DAO.Recordset
Dim qdfSig As DAO.QueryDef
Dim oldSQL As String
Dim newSQL As String
Dim i As Integer
Dim ctl
'If Parent.Name = "frmAFMDash" Then args = Nz(Parent.shiftList.Column(1), args)
'Me.RecordSource = "SELECT tblShiftManager.*, Format([shiftstart],'dd-mmm-yy') AS Expr1 FROM tblShiftManager WHERE (((Format([shiftstart],'dd-mmm-yy'))=Format('" & Nz(OpenArgs, args) & "','dd-mmm-yy')));"
    
    If IsNull(shiftCount) Then Exit Sub
'    If shiftCount > 6 Then
''        MsgBox "Only a maximum of 6 shifts can be loaded onto one report.", vbInformation, "AF3616"
'        DoCmd.close acReport, "new3616"
'    End If
    
    
    Set rsShift = CurrentDb.OpenRecordset(Me.RecordSource) 'The shifts
    Set qdfSig = CurrentDb.QueryDefs("qSig")
    oldSQL = qdfSig.sql
    
    i = 1
    With rsShift: Do While Not .EOF
    
        With Me.Controls("dSig" & i)
            '.SourceObject = "Report.dSig"
            
            If Not IsNull(rsShift!username) Then
                
                newSQL = oldSQL
                
                newSQL = Replace(newSQL, "[varUSN]", "'" & rsShift!username & "'")
                newSQL = Replace(newSQL, "[varSigTime]", "#" & rsShift!sigTime & "#")
                newSQL = Replace(newSQL, "[varShiftID]", rsShift!shiftID)
                
                .Report.RecordSource = newSQL
                '.Report.RecordSource = "SELECT UCase([tbluserauth].[lastname]) & '.' & UCase([tbluserauth].[firstname]) & '.' & Replace(UCase([@RANK].[rank]),'-','') & '.' & [tbluserauth].[username] AS sName, tblUserAuth.opInitials, tblUserAuth.username, tblShiftManager.shiftID, '" & rsShift!amosSigTime & "' AS sigTime FROM [@RANK] INNER JOIN (tblShiftManager INNER JOIN tblUserAuth ON tblShiftManager.amosSig = tblUserAuth.username) ON [@RANK].ID = tblUserAuth.rankID WHERE tblShiftManager.shiftID = " & rsShift!shiftID & ";"

               For Each ctl In .Report.Controls
                    ctl.Visible = True
                Next
            End If
            
        End With
        i = i + 1
        .MoveNext
    Loop: .MoveFirst: End With
    
    If Not rsShift.EOF Then
    
        If Not IsNull(rsShift!concur) Then
            'Me.namoSig.Report.RecordSource = "SELECT UCase([tbluserauth].[lastname]) & '.' & UCase([tbluserauth].[firstname]) & '.' & Replace(UCase([@RANK].[rank]),'-','') & '.' & [tbluserauth].[username] AS sName, tblUserAuth.opInitials, tblUserAuth.username, tblShiftManager.shiftID, '" & rsShift!namoSigTime & "' AS sigTime FROM [@RANK] INNER JOIN (tblShiftManager INNER JOIN tblUserAuth ON tblShiftManager.namoSig = tblUserAuth.username) ON [@RANK].ID = tblUserAuth.rankID WHERE tblShiftManager.shiftID = " & rsShift!shiftID & ";"
            Set qdfSig = CurrentDb.QueryDefs("qSig2")
            oldSQL = qdfSig.sql
            newSQL = oldSQL
            newSQL = Replace(newSQL, "[varUSN]", "'" & rsShift!concur & "'")
            newSQL = Replace(newSQL, "[varSigTime]", "#" & rsShift!concurSigTime & "#")
            newSQL = Replace(newSQL, "[varShiftID]", rsShift!shiftID)
            
            Me.namoSig.Report.RecordSource = newSQL
            
            Me.namoSig.Visible = True
            
            For Each ctl In Me.namoSig.Report.Controls
                ctl.Visible = True
            Next
        End If
        
        If Not IsNull(rsShift!cert) Then
            Set qdfSig = CurrentDb.QueryDefs("qSig3")
            oldSQL = qdfSig.sql
            newSQL = oldSQL
            newSQL = Replace(newSQL, "[varUSN]", "'" & rsShift!cert & "'")
            newSQL = Replace(newSQL, "[varSigTime]", "#" & rsShift!certSigTime & "#")
            newSQL = Replace(newSQL, "[varShiftID]", rsShift!shiftID)
            
            Me.afmSig.Report.RecordSource = newSQL
            
            Me.afmSig.Visible = True
            For Each ctl In Me.afmSig.Report.Controls
                ctl.Visible = True
            Next
            
        End If
    End If
    
End Sub

Private Sub Report_Open(cancel As Integer)
'OpenArgs = The DateSerial of the log
On Error GoTo sErr
args = Null

    Select Case Parent.Name
    
    Case "This throws an error when the parent doesn't exist"
        'error
    Case "frmAfmDash"
        args = Format(Nz(Parent.shiftList.Column(1), Nz(args, Date)), "dd-mmm-yy")
    Case "frmReportsView"
        args = Parent.Parent.reportDate
    End Select
    
    If Not IsNull(OpenArgs) Or Not IsNull(args) Then Me.RecordSource = "SELECT tblShiftManager.*, tblShiftCertifier.* FROM tblShiftCertifier " & _
                                                "RIGHT JOIN tblShiftManager ON tblShiftCertifier.ID = tblShiftManager.certifierID " & _
                                                "WHERE dateValue(tblShiftManager.shiftstart) = #" & Nz(Me.OpenArgs, Nz(args, Date)) & "#;"
sexit:
    Exit Sub
sErr:
    Select Case err
    Case 2452
        Resume Next
    Case Is <> 0
        ErrHandler err, Error$, Me.Name & ".Report_Open"
    End Select
End Sub

'Private Sub updateRecordSource()
'    Me.RecordSource = "SELECT tblShiftManager.*, Format([shiftstart],'dd-mmm-yy') AS Expr1 FROM tblShiftManager WHERE (((Format([shiftstart],'dd-mmm-yy'))=Format('" & Nz(OpenArgs, args) & "','dd-mmm-yy')));"
'End Sub

Private Sub savePDF_Click()
'Dim f As String
'Dim n As Integer
'Dim ans As VbMsgBoxResult
''log Me.RecordSource, Me.Name & ".savePDF_Click"
'    DoCmd.SetWarnings False
'    f = DLookup("drivePrefix", "tblSettings") & "1. Events Logs\" & UCase(Format(rDate, "yyyy\\mm mmm yyyy\\d mmm yy ") & "db") & ".pdf"
'    If dir(f) = "" Then createPath DLookup("drivePrefix", "tblSettings") & "1. Events Logs\" & UCase(Format(rDate, "yyyy\\mm mmm yyyy"))
'    Do While Len(dir(f)) > 0
'        n = n + 1
'        ans = IIf(ans <> 0, ans, MsgBox("A duplicate log was found for this date. replace?", vbQuestion + vbYesNoCancel, "Events Log"))
'        'This might not be right...
'        Select Case ans
'            Case vbYes
'                Exit Do
'            Case vbCancel
'                Exit Sub
'            Case vbNo
'                Select Case n
'                    Case 1
'                        f = Replace(f, ".pdf", " (" & n & ").pdf")
'                    Case Else
'                        f = Replace(f, " (" & n - 1 & ").pdf", " (" & n & ").pdf")
'                End Select
'                MsgBox "This log will be saved as '(" & n & ").pdf' instead", vbInformation, "AF3616"
'                Exit Do
'        End Select
'    Loop
''    If Not n = 0 Then
''        If MsgBox("A duplicate log was found for this date. replace?", vbQuestion + vbYesNo, "Events Log") = vbYes Then
''            MsgBox "A log with the same title was found. This log was saved as '(" & n & ").pdf' instead", vbInformation, "AF3616"
''        Else
'    'DoCmd.OpenReport "new3616", acViewReport, , Left(Mid(Me.RecordSource, InStr(1, Me.RecordSource, "WHERE ") + 6), Len(Mid(Me.RecordSource, InStr(1, Me.RecordSource, "WHERE ") + 6)) - 1)
'    DoCmd.OpenReport "new3616", acViewReport, , , , rDate
'    DoEvents
'    DoCmd.OutputTo acOutputReport, "new3616", acFormatPDF, f
'    'DoCmd.close acReport, "new3616", acSaveNo
'
'    DoCmd.SetWarnings True
'    'MsgBox "To sign, go to: " & vbCrLf & "Tools -> Certificates -> Digitally Sign" & vbCrLf & vbCrLf & "Then draw a signature block under 'Signature'" & vbCrLf & "(This log is not valid without a signature)", vbInformation, "AeroStat"
'    Select Case MsgBox("Saved successfully in " & f & "." & vbCrLf & "Open PDF?", vbQuestion + vbYesNo, "Events Log")
'        Case vbYes
'            Application.FollowHyperlink f
'            DoCmd.Close acReport, "new3616"
'
'    End Select
    Application.FollowHyperlink UtilForm3616.savePDF(rDate)
End Sub

Private Sub savePDF2_Click()
savePDF_Click
End Sub
