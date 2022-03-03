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
Option Explicit
Dim shiftID As Variant

Public Function getClipboard() As String
    On Error Resume Next
    Dim dataObj As New MSForms.DataObject
    dataObj.GetFromClipboard
    getClipboard = dataObj.GetText
    Set dataObj = Nothing
End Function

Private Sub bParse_Click()
Dim n As Integer
Dim logTime As Date
'If Nz(parse) = "" Then Exit Sub
If Nz(parse) = "" Then
    parse = getClipboard
End If
If Nz(parse) = "" Then Exit Sub

n = parseNOTAM(parse, DLookup("opinitials", "tbluserauth", "username='" & Util.getUser & "'"), Date, DateAdd("d", 3, Date))

    If n <> 0 Then
        Me.Filter = "ID = " & n
        Me.FilterOn = True
        
        Dim log As DAO.Recordset
'        Dim shiftID As Integer
'        shiftID = DLookup("lastshift", "tbluserauth", "username='" & Util.getUser & "'")
        'shiftID = IIf(     Nz(    DLookup("closed", "tblShiftManager", "shiftID = " & Nz(shiftID, 0))    , True), 0, shiftID)
        
        If logentry And shiftID <> 0 Then
            Set log = CurrentDb.OpenRecordset("SELECT * FROM tbl3616 WHERE shiftID = " & shiftID)
            With log
                Dim shiftStart As Date: shiftStart = CDate(Format(DLookup("shiftstart", "tblshiftmanager", "shiftid = " & shiftID), "dd-mmm-yy"))
                .AddNew
                !shiftID = shiftID
                
                'logTime = LToZ(Now)
                '!entryTime = IIf(IsNull(DLookup("entryTime", "tbl3616", "format(entryTime,'yymmddhhnn') = " & Format(logTime, "yymmddhhnn"))), logTime, DateAdd("n", 1, logTime))
                
                !entryTime = LToZ(shiftStart + TimeSerial(Hour(Now), Minute(Now), 0))
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
                !Initials = DLookup("opinitials", "tbluserauth", "username = '" & Util.getUser & "'")
                !entry = UCase(!entry)
                .update
                .Close
            End With
            Set log = Nothing
            

        End If
        
        MsgBox "Parse successful!", vbInformation, "NOTAM Control"
        DoCmd.Close acForm, "frmNOTAMParse"
        
'        If MsgBox("Parse successful, would you like to view the NOTAM?", vbQuestion + vbYesNo, "NOTAM Control") = vbYes Then
'            DoCmd.OpenForm "frmNOTAMInput", , , "ID = " & n, acFormEdit
'        End If
        
    End If
End Sub

Private Sub btnExample_Click()
MsgBox "Example of an ICAO Format:" & vbCrLf & "M0380/20 NOTAMN Q) ZMA/QFAGD/IV/NBO/A/000/999/2529N08023W005 A) KHST B) 2006291549 C) 2009262359 E) AERODROME OFFICIAL BUSINESS ONLY"
End Sub

Private Sub Form_Load()
    logentry.Visible = (Nz(Me.openargs) <> "")
    
    Select Case Nz(openargs)
        Case Is <> 0
            logentry.Visible = True
            shiftID = Me.openargs
            If Not Nz(DLookup("closed", "tblshiftmanager", "shiftid = " & shiftID), True) Then
                bParse_Click
            End If
    End Select



'If Nz(Me.OpenArgs) <> "" Then
'    logentry.Visible = True
'    shiftID = Me.OpenArgs
'End If

End Sub
