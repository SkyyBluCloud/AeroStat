Attribute VB_Name = "NOTAMUtil"
Option Compare Database
Private Const myName = "NOTAMUtil"

Public Function getDateFromNOTAM(ByVal s As String) As Date
If Nz(s) = "" Then Exit Function
getDateFromNOTAM = DateSerial(Left(s, 2), Mid(s, 3, 2), Mid(s, 5, 2)) & " " & TimeSerial(Left(Right(s, 4), 2), Right(s, 2), 0)
'getDate = getDate & " " & Left(Right(s, 4), 2) & ":" & Right(s, s)
End Function

Public Function convertNOTAMReport(Optional ByVal reportTable As String = "tblNOTAMReport")
Dim old As DAO.Recordset
Dim num As Integer
Set old = CurrentDb.OpenRecordset(reportTable)
If old.EOF Then Exit Function
log "Converting Report....", myName & ".convertNOTAMReport"

DoCmd.OpenForm "frmLoading"
old.MoveLast
old.MoveFirst
With Forms!frmLoading
    !loadingText.Caption = "Converting report..."
    DoEvents
    !pBar.Max = old.RecordCount
End With

With old: Do While Not .EOF
    num = num + 1
    parseNOTAM ![ICAO format], Nz(![NOTAM Originator]), ![Start Date], ![End Date], True
    .MoveNext
    If num Mod 20 = 0 Then log num & " records complete...", myName & ".convertNOTAMReport"
    DoEvents
    Forms!frmLoading!pBar.Value = Forms!frmLoading!pBar.Value + 1
    Loop
End With
log "Done!", myName & ".convertNTOAMReport"
DoCmd.Close acForm, "frmLoading"
End Function

Public Function retroCancel()
Dim N As DAO.Recordset
Dim cNOTAM As DAO.Recordset
Set cNOTAM = CurrentDb.OpenRecordset("SELECT * FROM tblNOTAM WHERE nType = 'C'")
With cNOTAM: Do While Not .EOF
    Set N = CurrentDb.OpenRecordset("SELECT * FROM tblNOTAM WHERE NOTAM = '" & !NOTAM & "'")
    If Not N.EOF Then
    With N:
        .edit
        !cancelled = True
        .Update
    End With
    End If
    .MoveNext
    Loop
End With
    
End Function

Public Function cancelNOTAM(ByVal cNOTAM As String, ByVal nEndTime As Date) As Boolean
On Error GoTo errtrap
Dim rNOTAM As DAO.Recordset
    If Nz(cNOTAM) <> "" Then
        Set rNOTAM = CurrentDb.OpenRecordset("SELECT * FROM tblNOTAM WHERE NOTAM = '" & cNOTAM & "'")
        With rNOTAM
            If Not .EOF Then
                .edit
                !isCancelled = True
                !endTime = nEndTime
                .Update
                .Close
            End If
        End With
    End If
fExit:
    cancelNOTAM = True
    Exit Function
    
errtrap:
    ErrHandler err, Error$, "NOTAMUtil.cancelNOTAM"
End Function

Public Function parseNOTAM(ByVal s As String, Optional ByVal usr As String, Optional ByVal start As String, Optional ByVal expiry As String, Optional ByVal proc As Boolean) As Integer
'Parses NOTAM in *ICAO FORMAT* and adds data to database NOTAM table.
On Error GoTo errtrap
parseNOTAM = 0
Dim N As String
Dim q, a, b, c, d, e As String
Dim cNOTAM As String
Dim cNOTAMStart As Date
Dim rNOTAM2 As DAO.Recordset
s = noBreaks(s)
Dim rNOTAM As DAO.Recordset
Set rNOTAM = CurrentDb.OpenRecordset("tblNOTAM")
With rNOTAM
    
    q = InStr(1, s, "Q)") + 3
    a = InStr(1, s, "A)") + 3
    b = InStr(1, s, "B)") + 3
    c = InStr(1, s, "C)") + 3
    d = InStr(1, s, "D)") + 3
    e = InStr(1, s, "E)") + 3
    
    .AddNew
    !NOTAM = Left(s, 8)
    If Not IsNull(DLookup("notam", "tblnotam", "notam = '" & !NOTAM & "'")) And Not proc Then
        MsgBox "This NOTAM already exists.", vbInformation, "NOTAM Control"
        Exit Function
    End If
    
    !nType = Mid(s, 15, 1)
    N = !nType 'Why?
    !aerodrome = Mid(s, a, 4)
    Select Case !nType
        Case "N", "R"
            !countryCode = Mid(s, q, (InStr(q, s, "/") - q))
            !qcode = Mid(s, InStr(q, s, "/") + 1, 5)
            
            !startTime = getDateFromNOTAM(Mid(s, b, 10))
            !endTime = getDateFromNOTAM(Mid(s, c, 10))
            If d > 3 Then
                !period = Mid(s, d, e - d - 4)
            End If
            If !nType = "R" Then
                cNOTAM = Mid(s, InStr(1, s, "NOTAMR") + 7, 8)
                !verbiage = "Replaces " & cNOTAM & ": " & Mid(s, e)
            Else
                !verbiage = Mid(s, e)
            End If
        Case "C"
            cNOTAM = Mid(s, 17, 8)
            !verbiage = "Cancel " & cNOTAM
            !startTime = IIf(Nz(start) = "", Now, start)
            !endTime = IIf(Nz(expiry) = "", DateAdd("d", 3, Now), expiry)
            cNOTAMStart = !startTime
        Case Else
            parseNOTAM = 0
            MsgBox "Could not parse NOTAM: Invalid format.", vbInformation, "NOTAM Control"
            Exit Function
    End Select
    
    If Nz(usr) <> "" Then
        If Len(usr) = 2 Then
            !issuedBy = usr
        ElseIf Not IsNull(DLookup("opinitials", "tbluserauth", "lastname = '" & Trim(Mid(usr, InStr(1, Trim(usr), " "))) & "'")) Then
            !issuedBy = DLookup("opinitials", "tbluserauth", "lastname = '" & Trim(Mid(usr, InStr(1, Trim(usr), " "))) & "'")
        ElseIf Nz(usr) <> "" Then
            !issuedBy = UCase(Left(usr, 1) & Mid(usr, InStr(1, Trim(usr), " ") + 1, 1)) & "*"
        End If
    End If
        
    .Update
    .Bookmark = .LastModified
    parseNOTAM = !ID
    
    If Nz(cNOTAM) <> "" Then
        Set rNOTAM2 = CurrentDb.OpenRecordset("SELECT * FROM tblNOTAM WHERE NOTAM = '" & cNOTAM & "'")
        With rNOTAM2
            .edit
            !isCancelled = True
            !endTime = Nz(cNOTAMStart, Now)
            .Update
            .Close
        End With
        .Close
    End If
End With
Set rNOTAM = Nothing
Set rNOTAM2 = Nothing
fExit:
    Exit Function
    Resume Next
errtrap:
    If err <> 3022 Then ErrHandler err, Error$, myName & ".parseNOTAM"
    
End Function
