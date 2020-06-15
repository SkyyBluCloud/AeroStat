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
fExit:
    log "Saved " & progress & "/" & total & " logs successfully!", "UtilForm3616.saveAllPDFs"
    saveAllPDFs = True
    Exit Function
errtrap:

End Function

Public Function savePDF(ByVal rDate As Date, Optional ByRef overwrite As Variant) As Boolean
On Error GoTo errtrap
Dim fp, f As String
Dim N As Integer
Dim ans As VbMsgBoxResult

    If Not IsMissing(overwrite) Then 'Triggeres if this method was called by another method (such as .saveAllPDFs)
        Select Case overwrite
        Case True
            ans = vbYes
        Case False
            ans = vbNo
        End Select
    End If

    DoCmd.SetWarnings False
    
    fp = DLookup("drivePrefix", "tblSettings") & UCase(Format(rDate, "yyyy\\mm mmm yy\\d mmm yy\\"))
    If dir(fp, vbDirectory) = "" Then createPath fp
    f = fp & UCase(Format(rDate, "d mmm yy ")) & "EVENTS LOG DB.PDF"
    Do While Len(dir(f)) > 0
        N = N + 1
        If ans = 0 Then
            ans = MsgBox("A duplicate log was found for this date. replace?", vbQuestion + vbYesNoCancel, "Events Log")
        End If

        Select Case ans
        Case vbYes
            Exit Do
        Case vbCancel
            Exit Function
        Case vbNo
            Select Case N
                Case 1
                    f = Replace(f, ".pdf", " (" & N & ").pdf")
                Case Else
                    f = Replace(f, " (" & N - 1 & ").pdf", " (" & N & ").pdf")
            End Select
        End Select
    Loop
    
    If N <> 0 And IsNull(overwrite) Then
        MsgBox "This log will be saved as '(" & N & ").pdf' instead", vbInformation, "AF3616"
    End If
    
    DoCmd.OpenReport "new3616", acViewReport, , , , rDate
    Reports!new3616.Visible = False
    DoEvents
    DoCmd.OutputTo acOutputReport, "new3616", acFormatPDF, f
    DoCmd.Close acReport, "new3616", acSaveNo
    
    DoCmd.SetWarnings True
    
    If IsNull(overwrite) Then
        Select Case MsgBox("Saved successfully in " & f & "." & vbCrLf & "Open PDF?", vbQuestion + vbYesNo, "Events Log")
        Case vbYes
            Application.FollowHyperlink f
            'DoCmd.Close acReport, "new3616"
        End Select
    End If
fExit:
    log "Successfully saved log as " & f, "UtilForm3616.savePDF"
    savePDF = True
    Exit Function
errtrap:
    ErrHandler err, Error$, "UtilForm3616.savePDF"
End Function

Public Function newEntry(ByVal shiftID As Integer, ByVal zuluDateTime As Date, ByVal entry As String, Optional ByVal opInitials As String) As Boolean
On Error GoTo errtrap
Dim dupeEntry As String
dupeEntry = Nz(DLookup("entry", "tbl3616", "shiftid = " & shiftID & " AND format(entrytime,'hhnn') = '" & Format(zuluDateTime, "hhnn" & "'")))
If dupeEntry <> "" Then
    If MsgBox("The following entry will be replaced:" & vbCrLf & Format(zuluDateTime, "hhnn") & ": " & dupeEntry & vbCrLf & vbCrLf & "Replace?", vbQuestion + vbYesNo, "Events Log") = vbNo Then
        While Not IsNull(DLookup("entrytime", "tbl3616", "shiftid = " & shiftID & " AND format(entrytime,'hhnn') = '" & Format(zuluDateTime, "hhnn" & "'")))
            zuluDateTime = DateAdd("n", 1, zuluDateTime)
        Wend
    Else
        CurrentDb.Execute "DELETE FROM tbl3616 WHERE entry = " & """" & dupeEntry & """", dbFailOnError
    End If
End If

If Nz(opInitials) = "" Then opInitials = Util.getOpInitials
entry = UCase(Trim(entry))

    CurrentDb.Execute "INSERT INTO tbl3616 (shiftID,entryTime,entry,initials) " & _
                        "SELECT " & shiftID & ", '" & Format(zuluDateTime, "dd-mmm-yy") & " " & _
                        Left(Format(zuluDateTime, "hhnn"), 2) & "." & Right(Format(zuluDateTime, "hhnn"), 2) & "', " & """" & entry & """" & ", '" & opInitials & "'", dbFailOnError

fExit:
    newEntry = True
    Exit Function
errtrap:
    MsgBox "Entry could not be made. (" & err & ")", vbCritical, "Events Log"
    handleError "newEntry"
End Function

Public Function signLog(ByVal shiftID As Integer, ByVal eLogRecSrc As String, ByVal role As Integer) As Boolean
On Error GoTo errtrap
Dim rs As DAO.Recordset
Set rs = CurrentDb.OpenRecordset(eLogRecSrc)
Dim roleStr As String
If rs.RecordCount = 0 Then Exit Function
'If rs!closed Then
'    MsgBox "This shift was already signed.", vbInformation, "Events Log"
'    Exit Function
'End If
    While shiftID <> rs!shiftID
        rs.MoveNext
    Wend

    Select Case role
    Case 2
        If rs!closed Then
            MsgBox "This shift was already signed.", vbInformation, "Events Log"
            Exit Function
        End If
        
        roleStr = "AMOS"
        
    Case 3: roleStr = "NAMO"
    Case 4: roleStr = "AFM"
    End Select
    
    If MsgBox("You are signing this Events Log as the " & roleStr & ". " & vbCrLf & vbCrLf & _
        "By signing this document, you certify that all entries are correct; " & _
        "that all scheduled operations have been accomplished, except as noted; " & _
        "that all abnormal occurences or conditions and all significant incidents/events have been recorded.", vbOKCancel + vbInformation, "Events Log") = vbCancel _
    Then Exit Function
    
    With rs
        If role = 2 Then
            .edit
            !closed = True
            .Fields(LCase(roleStr) & "Sig") = getUSN
            .Fields(LCase(roleStr) & "SigTime") = Now
            .Update
        Else
            CurrentDb.Execute "UPDATE tblShiftManager SET " & LCase(roleStr) & "Sig = '" & getUSN & "', " & _
                                                                LCase(roleStr) & "SigTime = Now() " & _
                                                                Mid(eLogRecSrc, InStr(1, eLogRecSrc, "WHERE"), Len(eLogRecSrc))
        End If
    End With
    
fExit:
    MsgBox "Log signed!", vbInformation, "Events Log"
    signLog = True
    Exit Function
errtrap:
    MsgBox "The log was NOT signed." & vbCrLf & "(" & err & ")", vbCritical, "Events Log"
    handleError "signLog"
End Function
