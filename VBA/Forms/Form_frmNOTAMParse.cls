VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmNOTAMParse"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Dim shiftID As Integer

Private Sub bParse_Click()
Dim N As Integer
Dim logTime As Date
'If Nz(parse) = "" Then Exit Sub
If Nz(parse) = "" Then
    Dim DataObj As New MSForms.DataObject
    DataObj.GetFromClipboard
    parse = DataObj.GetText
End If
If Nz(parse) = "" Then Exit Sub

N = parseNOTAM(parse, DLookup("opinitials", "tbluserauth", "username='" & Environ$("username") & "'"), Date, DateAdd("d", 3, Date))

    If N <> 0 Then
        Me.filter = "ID = " & N
        Me.FilterOn = True
        
        Dim log As DAO.Recordset
'        Dim shiftID As Integer
'        shiftID = DLookup("lastshift", "tbluserauth", "username='" & Environ$("username") & "'")
        shiftID = IIf(Nz(DLookup("closed", "tblShiftManager", "shiftID =" & shiftID), True), 0, shiftID)
        
        If logentry And shiftID <> 0 Then
            Set log = CurrentDb.OpenRecordset("SELECT * FROM tbl3616 WHERE shiftID = " & shiftID)
            With log
                .AddNew
                !shiftID = shiftID
                'logTime = LToZ(Now)
                '!entryTime = IIf(IsNull(DLookup("entryTime", "tbl3616", "format(entryTime,'yymmddhhnn') = " & Format(logTime, "yymmddhhnn"))), logTime, DateAdd("n", 1, logTime))
                !entryTime = LToZ(Now)
                Do While Not IsNull(DLookup("entryTime", "tbl3616", "format(entryTime,'yymmddhhnn') = " & Format(!entryTime, "yymmddhhnn")))
                    !entryTime = DateAdd("n", 1, !entryTime)
                Loop
                
                Select Case nType
                    Case "N", "R"
                        !entry = "NEW NOTAM: " & NOTAM & " - " & verbiage & ". VALID " & Format(startTime, "yymmddhhnn") & " - " & Format(endTime, "yymmddhhnn") & ". " & _
                            "NEW NOTAM COUNT: " & DCount("notam", "qNOTAMM") & "M, " & DCount("notam", "qNOTAML") & "L"
                    Case "C"
                        !entry = "NOTAM " & NOTAM & " CANCELS " & Right(verbiage, 8) & ": " & DLookup("verbiage", "tblnotam", "notam = '" & Right(verbiage, 8) & "'") & ". " & _
                            "NEW NOTAM COUNT: " & DCount("notam", "qNOTAMM") & "M, " & DCount("notam", "qNOTAML") & "L"
                End Select
                !Initials = DLookup("opinitials", "tbluserauth", "username = '" & Environ$("username") & "'")
                .update
                .Close
            End With
            Set log = Nothing
        End If
        
        MsgBox "Parse successful!", vbInformation, "NOTAM Control"
        DoCmd.Close acForm, "frmNOTAMParse", acSaveYes
        
'        If MsgBox("Parse successful, would you like to view the NOTAM?", vbQuestion + vbYesNo, "NOTAM Control") = vbYes Then
'            DoCmd.OpenForm "frmNOTAMInput", , , "ID = " & n, acFormEdit
'        End If
        
    End If
End Sub

Private Sub btnExample_Click()
MsgBox "Example of an ICAO Format:" & vbCrLf & "M0380/20 NOTAMN Q) ZMA/QFAGD/IV/NBO/A/000/999/2529N08023W005 A) KHST B) 2006291549 C) 2009262359 E) AERODROME OFFICIAL BUSINESS ONLY"
End Sub

Private Sub Form_Load()
If Nz(Me.OpenArgs) <> "" Then
    logentry.Visible = True
    shiftID = Me.OpenArgs
End If
logentry.Visible = (Nz(Me.OpenArgs) <> "")
End Sub
