Attribute VB_Name = "UtilForm3616"
'Air Force Form 3616
Option Compare Database

Private Sub handleError(ByVal caller As String)
errHandler err, Error$, "UtilForm3616." & caller
End Sub

Public Function newEntry(ByVal shiftID As Integer, ByVal zuluDateTime As Date, ByVal entry As String, Optional ByVal opInitials As String) As Boolean
On Error GoTo errTrap
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
errTrap:
    MsgBox "Entry could not be made. (" & err & ")", vbCritical, "Events Log"
    handleError "newEntry"
End Function

Public Function signLog(ByVal shiftID As Integer, ByVal eLogRecSrc As String, ByVal role As Integer) As Boolean
On Error GoTo errTrap
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
            .update
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
errTrap:
    MsgBox "The log was NOT signed." & vbCrLf & "(" & err & ")", vbCritical, "Events Log"
    handleError "signLog"
End Function
