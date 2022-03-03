Attribute VB_Name = "UtilForm3616"
'Air Force Form 3616
Option Compare Database

Private Sub handleError(ByVal caller As String)
ErrHandler err, Error$, "UtilForm3616." & caller
End Sub

Public Function saveAllPDFs(ByVal startDate As Date, Optional ByVal endDate As Date, Optional ByVal overwrite As Boolean = True) As Boolean
On Error Resume Next
endDate = Nz(endDate, Date)
Dim d, progress, total As Integer
total = (DateDiff("d", startDate, endDate)) + 1

    log "Attempting to save " & total & " logs...", "UtilForm3616.saveAllPDFs"
    For d = startDate To endDate
        If Not savePDF(d, overwrite) Then
            log "Could not save this log; let's try the next one!", "UtilForm3616.saveAllPDFs", "WARN"
        Else
            progress = progress + 1
        End If
    Next
fexit:
    log "Saved " & progress & "/" & total & " logs successfully!", "UtilForm3616.saveAllPDFs"
    saveAllPDFs = True
    Exit Function
errtrap:

End Function

Public Function savePDF(ByVal rDate As Date, Optional ByRef overwrite As Variant) As String
On Error GoTo errtrap
Dim fp, f As String
Dim n As Integer
Dim ans As VbMsgBoxResult

    If Not IsMissing(overwrite) Then 'Triggers if this method was called by another method (such as .saveAllPDFs)
        Select Case overwrite
        Case True
            ans = vbYes
        Case False
            ans = vbNo
        End Select
    End If

    'DoCmd.SetWarnings False
    If IsNull(DLookup("data", "tblSettings", "key = 'dbRoot'")) Then GoTo errtrap
    
    fp = DLookup("data", "tblSettings", "key = 'dbRoot'") & "Daily Operations\" & UCase(Format(rDate, "yyyy\\mm mmm yy\\dd\\"))
    If dir(fp, vbDirectory) = "" Then createPath fp
    f = fp & UCase(Format(rDate, "d mmm yy ")) & "EVENTS LOG DB.PDF"
    
    Do While Len(dir(f)) > 0
        n = n + 1
        If ans = 0 Then
            ans = MsgBox("A duplicate log was found for this date. replace?", vbQuestion + vbYesNoCancel, "Events Log")
        End If

        Select Case ans
        Case vbYes
            Exit Do
        Case vbCancel
            Exit Function
        Case vbNo
            Select Case n
                Case 1
                    f = Replace(f, ".pdf", " (" & n & ").pdf")
                Case Else
                    f = Replace(f, " (" & n - 1 & ").pdf", " (" & n & ").pdf")
            End Select
        End Select
    Loop
    
    If n <> 0 And IsNull(overwrite) Then
        MsgBox "This log will be saved as '(" & n & ").pdf' instead", vbInformation, "AF3616"
    End If
    
    DoCmd.OpenReport "new3616", acViewReport, , , acHidden, rDate
    DoEvents
    
    DoCmd.OutputTo acOutputReport, "new3616", acFormatPDF, f
    DoCmd.Close acReport, "new3616", acSaveNo
    
    'DoCmd.SetWarnings True
    
    If IsNull(overwrite) Then
        Select Case MsgBox("Saved successfully in " & f & "." & vbCrLf & "Open PDF?", vbQuestion + vbYesNo, "Events Log")
        Case vbYes
            Application.FollowHyperlink f
            'DoCmd.Close acReport, "new3616"
        End Select
    End If
fexit:
        

    
    log "Successfully saved log as " & f, "UtilForm3616.savePDF"
    savePDF = f
    MsgBox "Log saved.", vbInformation, "Save PDF"
    Exit Function
errtrap:
    MsgBox "The log could not be saved. (" & err & ").", vbCritical, "AF3616"
    ErrHandler err, Error$, "UtilForm3616.savePDF"
End Function

Public Function newEntry(ByVal shiftID As Integer, ByVal zuluDateTime As Date, ByVal entry As String, Optional ByVal opInitials As String) As Boolean
On Error GoTo errtrap
Dim dupeEntry As String
dupeEntry = Nz(DLookup("entry", "tbl3616", "shiftid = " & shiftID & " AND Not archive AND entrytime = #" & zuluDateTime & "#"))

    If dupeEntry <> "" And Not Util.getSettings("frm3616AllowDuplicateTimes") Then
        If MsgBox("The following entry will be replaced:" & vbCrLf & Format(zuluDateTime, "hhnn") & ": " & dupeEntry & vbCrLf & vbCrLf & "Replace?", vbQuestion + vbYesNo, "Events Log") = vbNo Then
            While Not IsNull(DLookup("entrytime", "tbl3616", "shiftid = " & shiftID & " AND format(entrytime,'hhnn') = '" & Format(zuluDateTime, "hhnn" & "'")))
                zuluDateTime = DateAdd("n", 1, zuluDateTime)
            Wend
        Else
            'CurrentDb.Execute "UPDATE tbl3616 SET archive = True, archiveBy = '" & opInitials & "' WHERE entry = " & """" & dupeEntry & """", dbFailOnError
            Dim db As DAO.Database: Set db = CurrentDb
            db.Execute "UPDATE tbl3616 SET archive = True, archiveBy = '" & opInitials & "', archiveTime = Now() WHERE entrytime = #" & zuluDateTime & "# AND entry = " & """" & dupeEntry & """", dbFailOnError
            log "Archived " & db.RecordsAffected, "UtilForm3616.newEntry"
            Set db = Nothing
        End If
    End If
    
    If Nz(opInitials) = "" Then opInitials = Util.getOpInitials(Util.getUser)
    entry = UCase(Trim(entry))

    CurrentDb.Execute "INSERT INTO tbl3616 (shiftID,originalOpInitials,entryTime,entry,initials) " & _
                        "SELECT " & shiftID & ", getopinitials(getuser()), '" & Format(zuluDateTime, "dd-mmm-yy") & " " & _
                        Left(Format(zuluDateTime, "hhnn"), 2) & "." & Right(Format(zuluDateTime, "hhnn"), 2) & "', " & """" & entry & """" & ", '" & opInitials & "'", dbFailOnError

fexit:
    newEntry = True
    Exit Function
errtrap:
    MsgBox "Entry could not be made. (" & err & ")", vbCritical, "Events Log"
    ErrHandler err, Error$, "Util3616.newEntry"
    Exit Function
    Resume Next
End Function

Public Function signLog(ByVal shiftID As Integer, ByVal role As Integer) As Boolean
On Error GoTo errtrap
Dim db As DAO.Database: Set db = CurrentDb
Dim rsShift As DAO.Recordset: Set rsShift = CurrentDb.OpenRecordset("SELECT * FROM tblShiftManager WHERE shiftID = " & shiftID)
Dim roleStr As String
Dim cert As Variant
If rsShift.RecordCount = 0 Then Exit Function
'If rs!closed Then
'    MsgBox "This shift was already signed.", vbInformation, "Events Log"
'    Exit Function
'End If
    
    Select Case role
    Case 1
        If rsShift!certifierID <> 0 Then
        'If IsNull(amosSig) Then
            MsgBox "This shift was already signed.", vbInformation, "Events Log"
            Exit Function
        End If
        
        roleStr = "AMOS"
        
    Case 2
        If rsShift!certifierID <> 0 Then
            MsgBox "This shift was already signed.", vbInformation, "Events Log"
            Exit Function
        End If
        
        roleStr = "NAMO"
    Case 3
        If rsShift!certifierID <> 0 Then
            MsgBox "This shift was already signed.", vbInformation, "Events Log"
            Exit Function
        End If
        roleStr = "AFM"
    End Select
    
    'TODO: Don't hard-code
    If MsgBox("You are signing as the " & roleStr & ". " & vbCrLf & vbCrLf & _
        "By signing this document, you certify that all entries are correct; " & _
        "that all scheduled operations have been accomplished, except as noted; " & _
        "that all abnormal occurences or conditions and all significant incidents/events have been recorded.", vbOKCancel + vbInformation, "Events Log") = vbCancel _
    Then Exit Function
    
'    If Not IsNull(DLookup("initials", "tbl3616", "shiftID = " & shiftID & " AND right(initials,1) <> '*'")) Then
'        CurrentDb.Execute "UPDATE tblShiftManager SET reviewerComments = '* = Denotes entry re-accomplished. " & reviewerComments & "' WHERE shiftID = " & shiftID, dbFailOnError
'    End If
    
    With rsShift
        Select Case role
        Case 1
            .edit
            !closed = True
            !certifierID = newCert(Util.getUser)
            !amosSig = Util.getUser
            !amosSigTime = Now
            .update
        Case 2 Or 3
            If Not UtilCertifier.certifyShiftDay(role, Util.getUser, DateValue(!shiftStart)) Then GoTo errtrap
            
        End Select
            'Old Sig
'            CurrentDb.Execute "UPDATE tblShiftManager SET " & LCase(roleStr) & "Sig = '" & util.getuser & "', " & _
'                                                                LCase(roleStr) & "SigTime = Now() " & _
'                                                                Mid(eLogRecSrc, InStr(1, eLogRecSrc, "WHERE"), Len(eLogRecSrc))
    End With
    
fexit:
    MsgBox "Log signed!", vbInformation, "Events Log"
    signLog = True
    Exit Function
errtrap:
    MsgBox "The log was NOT signed." & vbCrLf & "(" & err & ")", vbCritical, "Events Log"
    handleError "signLog"
End Function
