Attribute VB_Name = "UtilCertifier"
Option Compare Database
Option Explicit

'Converts the old signature format into the Certifier format.
Sub transform()
On Error GoTo sErr
Dim db As DAO.Database: Set db = CurrentDb
Dim rsShift As DAO.Recordset: Set rsShift = db.OpenRecordset("SELECT tblShiftManager.*, tblUserAuth.username FROM tblShiftManager INNER JOIN tblUserAuth ON tblUserAuth.opInitials = Right(tblShiftManager.superLead,2) WHERE certifierID = 0 AND (amosSig Is Not Null)")

    log "Transforming " & rsShift.RecordCount & " signatures to Certifier format...", "UtilCertifier.transform"
    
    With rsShift: Do While Not .EOF
        .edit
        DoEvents
        
        !certifierID = newCert(!amosSig, !amosSigTime, !namoSig, !namoSigTime, !afmSig, !afmSigTime)
        .update
        DoEvents
        
        .MoveNext
        If .PercentPosition > 0 And .PercentPosition Mod 10 = 0 Then log .PercentPosition & "% completed...", "UtilCertifier.transform"
        DoEvents
        
    Loop: End With
    log "Done!", "UtilCertifier.transform"
sexit:
    Exit Sub
sErr:
    ErrHandler err, Error$, "UtilCertifier.transform"
    Stop
End Sub

'Creates a new certifier record; returns the cert number to be appended to its respective record/table
Public Function newCert(ByVal usn As String, Optional ByVal sigTime As Variant = Null, _
                        Optional ByVal concur As Variant = Null, Optional ByVal concurSigTime As Variant = Null, _
                        Optional ByVal certUSN As Variant = Null, Optional ByVal certSigTime As Variant = Null) As Double
On Error GoTo fErr

sigTime = Nz(sigTime, Now)
concurSigTime = Nz(concurSigTime, Now)
certSigTime = Nz(certSigTime, Now)

Dim db As DAO.Database: Set db = CurrentDb

Dim auth As Variant: auth = DLookup("authLevel", "tblUserAuth", "username = '" & usn & "'")

If auth > 6 Then GoTo noauth

    
    db.Execute "INSERT INTO tblShiftCertifier (username, sigTime, concur, concurSigTime, cert, certSigTime) SELECT '" & usn & "', #" & sigTime & "#" & _
                IIf(Not IsNull(concur), ", '" & concur & "', #" & concurSigTime & "#", ", Null, Null") & _
                IIf(Not IsNull(certUSN), ", '" & certUSN & "', #" & certSigTime & "#", ", Null, Null")
                
    newCert = db.OpenRecordset("SELECT @@IDENTITY").Fields(0)
    
fexit:
    Set db = Nothing
    log "New cert! (" & CStr(newCert) & ")", "UtilCertifier.newCert"
    'Set rs = Nothing
    Exit Function
noauth:
    log "You don't have permission to do this.", "UtilCertifier.newCert", "WARN"
    GoTo fexit
fErr:
    ErrHandler err, Error$, "UtilCertifier.newCert"
    
End Function

'Finds the existing cert; Returns success
Public Function deCertifyDay(ByVal level As Integer, ByVal usn As String, ByVal reportDate As Date) As Boolean
Dim sql As String
    Select Case level
        Case 2
            sql = "SET concur = Null, concurSigTime = Null"
        Case 3
            sql = "SET cert = Null, certSigTime = Null, concur = Null, concurSigTime = Null"
        Case Else: Exit Function
    End Select
    
    Dim db As DAO.Database: Set db = CurrentDb
    db.Execute "UPDATE tblShiftCertifier INNER JOIN tblShiftManager ON tblShiftManager.certifierID = tblShiftCertifier.ID " & sql & " WHERE DateValue(shiftStart) = #" & reportDate & "#"
    log "Removed " & db.RecordsAffected & " signatures. (Level " & level & ")", "UtilCertifier.deCertifyDay"
    
fexit:
    deCertifyDay = True
    Exit Function
fErr:
    ErrHandler err, Error$, "UtilCertifier.deCertifyDay"

End Function

'Finds all the certs for the day, and adds the additional signatures; Returns success
Public Function certifyShiftDay(ByVal level As Integer, ByVal usn As String, ByVal reportDate As Date) As Boolean
On Error GoTo fErr
Dim auth As Variant: auth = DLookup("authLevel", "tblUserAuth", "username = '" & usn & "'") ' Get/set authLevel
Dim reqAuth As Integer 'Required Authoritiy
Dim db As DAO.Database: Set db = CurrentDb
If IsNull(auth) Then Exit Function
If level < 2 Or level > 3 Then Exit Function 'Only accept numbers 2 - 3
If IsNull(DLookup("ID", "tblUserAuth", "username = '" & usn & "'")) Then Exit Function 'Check for existing user
    
    'Open certs based on the shift day.
    Dim RS As DAO.Recordset
    Set RS = db.OpenRecordset("SELECT * FROM tblShiftCertifier RIGHT JOIN tblShiftManager ON tblShiftManager.certifierID = tblShiftCertifier.ID " & _
                                                        "WHERE DateValue(shiftStart) = #" & reportDate & "#", , dbFailOnError)
    With RS
        If .EOF Then .AddNew
        Select Case level
        Case 2 'NAMO
            
            reqAuth = 3
            If auth > reqAuth Then GoTo noauth
            
            If Nz(!concur) <> "" Then
                MsgBox "This record has already been signed.", vbInformation, "Certifier"
                GoTo fexit
            End If
            
            Do While Not .EOF
                If Nz(!username) = "" Then
                    MsgBox "All shifts for this day have not been signed yet.", vbExclamation, "Certifier"
                    GoTo fexit
                End If
                .MoveNext
            Loop
            .MoveFirst
            
            'db.Execute "UPDATE tblShiftCertifier INNER JOIN tblShiftManager ON tblShiftManager.certifierID = tblShiftCertifier.ID " & _
                                "SET concur = '" & usn & "', concurSigTime = Now() WHERE DateValue(shiftstart) = #" & reportDate & "#", dbFailOnError
            !concur = usn
            !concurSigTime = Now
            
            
            
            Dim qdf As DAO.QueryDef: Set qdf = CurrentDb.QueryDefs("qTrafficCert")
            With qdf
                .Parameters("varCertifierID") = RS!certifierID
                .Parameters("varDate") = reportDate
                .Execute dbFailOnError
                log "Certified " & .RecordsAffected & " flight plans.", "UtilCertifier.certifyShiftDay"
            End With
            
        Case 3 'AFM
            reqAuth = 2
            If auth > reqAuth Then GoTo noauth
            
            If Nz(!cert) <> "" Then
                MsgBox "This record has already been signed.", vbInformation, "Certifier"
                GoTo fexit
            End If
            
            If Nz(!username) = "" Then
                MsgBox "This report has not been reviewed by the NAMO/AMOM yet.", vbExclamation, "Certifier"
                GoTo fexit
            End If
            db.Execute "UPDATE tblShiftCertifier INNER JOIN tblShiftManager ON tblShiftManager.certifierID = tblShiftCertifier.ID " & _
                                "SET cert = '" & usn & "', certSigTime = Now() WHERE DateValue(shiftstart) = #" & reportDate & "#", dbFailOnError

        End Select
        
    End With
    certifyShiftDay = True
    log "Certified for the day (" & reportDate & "); Level " & level & " signature applied to " & db.RecordsAffected & " certs.", "UtilCertifier.certifyShiftDay"
    
fexit:
    RS.Close
    Set RS = Nothing
    Exit Function
noauth:
    log "You don't have permission to do this." & vbCrLf & "(Expected level " & reqAuth & ", but returned " & level & ")", "UtilCertifier.certifyShiftDay", "WARN"
    GoTo fexit
fErr:
    ErrHandler err, Error$, "UtilCertifier.certifyShiftDay"
    Resume Next
End Function

