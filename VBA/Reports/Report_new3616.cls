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
Public args As String

Private Declare PtrSafe Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" ( _
                    ByVal hwnd As Long, _
                    ByVal lpOperation As String, _
                    ByVal lpFile As String, _
                    ByVal lpParameters As String, _
                    ByVal lpDirectory As String, _
                    ByVal nShowCmd As Long) As Long

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

authLevel = DLookup("authLevel", "tbluserauth", "username = '" & Environ$("username") & "'")
If authLevel > 4 Then Exit Sub

    'if msgbox("By signing this document, you certify that all entries are correct and the form contains all required entries as

End Sub

Private Sub Report_Load()
On Error Resume Next
Dim RS As DAO.Recordset
Dim i As Integer
Dim ctl
'If Parent.Name = "frmAFMDash" Then args = Nz(Parent.shiftList.Column(1), args)
'Me.RecordSource = "SELECT tblShiftManager.*, Format([shiftstart],'dd-mmm-yy') AS Expr1 FROM tblShiftManager WHERE (((Format([shiftstart],'dd-mmm-yy'))=Format('" & Nz(OpenArgs, args) & "','dd-mmm-yy')));"

    If IsNull(shiftCount) Then Exit Sub
'    If shiftCount > 6 Then
''        MsgBox "Only a maximum of 6 shifts can be loaded onto one report.", vbInformation, "AF3616"
'        DoCmd.close acReport, "new3616"
'    End If
    
    Set RS = CurrentDb.OpenRecordset(Me.RecordSource)
    RS.filter = "not IsNull(amosSig)"
    

    i = 1
    With RS: Do While Not .EOF
        With Me.Controls("dSig" & i)
            If Not IsNull(RS!amosSig) Then 'AMOS
            '.SourceObject = "Report.dSig"
                For Each ctl In .Controls
                    ctl.Visible = True
                Next
                .Report.RecordSource = "SELECT UCase([tbluserauth].[lastname]) & '.' & UCase([tbluserauth].[firstname]) & '.' & Replace(UCase([RANK].[rank]),'-','') & '.' & [tbluserauth].[username] AS sName, tblUserAuth.opInitials, tblUserAuth.username, tblShiftManager.shiftID, '" & RS!amosSigTime & "' AS sigTime FROM RANK INNER JOIN (tblShiftManager INNER JOIN tblUserAuth ON tblShiftManager.amosSig = tblUserAuth.username) ON RANK.ID = tblUserAuth.rankID WHERE tblShiftManager.shiftID = " & RS!shiftID & ";"
            End If
        End With
        i = i + 1
        .MoveNext
    Loop: .MoveFirst: End With
    
    
    If Not IsNull(RS!namoSig) Then
        Me.namoSig.Visible = True
        For Each ctl In Me.namoSig.Report.Controls
            ctl.Visible = True
        Next
        Me.namoSig.Report.RecordSource = "SELECT UCase([tbluserauth].[lastname]) & '.' & UCase([tbluserauth].[firstname]) & '.' & Replace(UCase([RANK].[rank]),'-','') & '.' & [tbluserauth].[username] AS sName, tblUserAuth.opInitials, tblUserAuth.username, tblShiftManager.shiftID, '" & RS!namoSigTime & "' AS sigTime FROM RANK INNER JOIN (tblShiftManager INNER JOIN tblUserAuth ON tblShiftManager.amosSig = tblUserAuth.username) ON RANK.ID = tblUserAuth.rankID WHERE tblShiftManager.shiftID = " & RS!shiftID & ";"
    End If
    
    If Not IsNull(RS!afmSig) Then
        Me.afmSig.Visible = True
        For Each ctl In Me.afmSig.Report.Controls
            ctl.Visible = True
        Next
        Me.afmSig.Report.RecordSource = "SELECT UCase([tbluserauth].[lastname]) & '.' & UCase([tbluserauth].[firstname]) & '.' & Replace(UCase([RANK].[rank]),'-','') & '.' & [tbluserauth].[username] AS sName, tblUserAuth.opInitials, tblUserAuth.username, tblShiftManager.shiftID, '" & RS!afmSigTime & "' AS sigTime FROM RANK INNER JOIN (tblShiftManager INNER JOIN tblUserAuth ON tblShiftManager.amosSig = tblUserAuth.username) ON RANK.ID = tblUserAuth.rankID WHERE tblShiftManager.shiftID = " & RS!shiftID & ";"
    End If
    
End Sub

Private Sub Report_Open(cancel As Integer)
On Error Resume Next
    If Parent.Name <> "frmAFMDash" Then
        'error
    Else
        args = Nz(Parent.shiftList.Column(1), args)
    End If
    If Not IsNull(OpenArgs) Or args <> "" Then Me.RecordSource = "SELECT tblShiftManager.*, Format([shiftstart],'dd-mmm-yy') AS Expr1 FROM tblShiftManager WHERE (((Format([shiftstart],'dd-mmm-yy'))=Format('" & Nz(OpenArgs, args) & "','dd-mmm-yy')));"
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
    If Not UtilForm3616.savePDF(rDate) Then Exit Sub
End Sub

Private Sub savePDF2_Click()
savePDF_Click
End Sub
