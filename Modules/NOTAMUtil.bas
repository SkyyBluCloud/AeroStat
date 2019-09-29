Attribute VB_Name = "NOTAMUtil"
Option Compare Database
Private Const myName = "NOTAMUtil"

Public Function getDateFromNOTAM(ByVal s As String) As Date
If Nz(s) = "" Then Exit Function
getDateFromNOTAM = DateSerial(Left(s, 2), Mid(s, 3, 2), Mid(s, 5, 2)) & " " & TimeSerial(Left(Right(s, 4), 2), Right(s, 2), 0)
'getDate = getDate & " " & Left(Right(s, 4), 2) & ":" & Right(s, s)
End Function

Public Function convertNOTAMReport(ByVal reportTable As String)
Dim old As DAO.Recordset
Dim num As Integer
Set old = CurrentDb.OpenRecordset(reportTable)
If old.EOF Then Exit Function
errHandler 0, "Converting Report....", myName
With old: Do While Not .EOF
    num = num + 1
    parseNOTAM ![ICAO format], ![NOTAM Originator], ![Start Date], ![End Date]
    .MoveNext
    If num Mod 20 = 0 Then errHandler 0, num & " records complete...", myName
    DoEvents
    Loop
End With
errHandler 0, "DONE!", myName
End Function

Public Function retroCancel()
Dim n As DAO.Recordset
Dim cNOTAM As DAO.Recordset
Set cNOTAM = CurrentDb.OpenRecordset("SELECT * FROM tblNOTAM WHERE nType = 'C'")
With cNOTAM: Do While Not .EOF
    Set n = CurrentDb.OpenRecordset("SELECT * FROM tblNOTAM WHERE NOTAM = '" & !NOTAM & "'")
    If Not n.EOF Then
    With n:
        .edit
        !cancelled = True
        .update
    End With
    End If
    .MoveNext
    Loop
End With
    
End Function

Public Function cancelNOTAM(ByVal cNOTAM As String, ByVal nEndTime As Date) As Boolean
On Error GoTo errTrap
Dim rNOTAM As DAO.Recordset
    If Nz(cNOTAM) <> "" Then
        Set rNOTAM = CurrentDb.OpenRecordset("SELECT * FROM tblNOTAM WHERE NOTAM = '" & cNOTAM & "'")
        With rNOTAM
            If Not .EOF Then
                .edit
                !isCancelled = True
                !endTime = nEndTime
                .update
                .Close
            End If
        End With
    End If
fExit:
    cancelNOTAM = True
    Exit Function
    
errTrap:
    errHandler err, Error$, "NOTAMUtil.cancelNOTAM"
End Function

Public Function parseNOTAM(ByVal s As String, Optional ByVal usr As String, Optional ByVal start As String, Optional ByVal expiry As String) As Integer
'Parses NOTAM in *ICAO FORMAT* and adds data to database NOTAM table.
On Error GoTo errTrap
parseNOTAM = 0
Dim n As String
Dim cNOTAM As String
Dim cNOTAMStart As Date
Dim rNOTAM2 As DAO.Recordset
s = noBreaks(s)
Dim rNOTAM As DAO.Recordset
Set rNOTAM = CurrentDb.OpenRecordset("tblNOTAM")
With rNOTAM
    
    q = InStr(1, s, " Q)") + 4
    a = InStr(1, s, "A)") + 3
    b = InStr(1, s, " B)") + 4
    c = InStr(1, s, " C)") + 4
    d = InStr(1, s, " D)") + 4
    e = InStr(1, s, " E)") + 4
    
    .AddNew
    !NOTAM = Left(s, 8)
    If Not IsNull(DLookup("notam", "tblnotam", "notam = '" & !NOTAM & "'")) Then
        parseNOTAM = 0
        MsgBox "This NOTAM already exists.", vbInformation, "NOTAM Control"
        Exit Function
    End If
    
    !nType = Mid(s, 15, 1)
    n = !nType 'Why?
    !aerodrome = Mid(s, a, 4)
    Select Case !nType
        Case "N", "R"
            !countryCode = Mid(s, q, 4)
            !qcode = Mid(s, q + 5, 5)
            
            !startTime = getDateFromNOTAM(Mid(s, b, 10))
            !endTime = getDateFromNOTAM(Mid(s, c, 10))
            If d > 4 Then
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
        
    .update
    .Bookmark = .LastModified
    parseNOTAM = !ID
    
    If Nz(cNOTAM) <> "" Then
        Set rNOTAM2 = CurrentDb.OpenRecordset("SELECT * FROM tblNOTAM WHERE NOTAM = '" & cNOTAM & "'")
        With rNOTAM2
            .edit
            !isCancelled = True
            !endTime = Nz(cNOTAMStart, Now)
            .update
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
errTrap:
    errHandler err, Error$, myName & ".parseNOTAM"
    
End Function
