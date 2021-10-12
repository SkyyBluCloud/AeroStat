VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmDupeView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Dim rsDUPE As DAO.Recordset

Dim RS As DAO.Recordset
Dim qdf As DAO.QueryDef

Private Sub btnAll_Click()
If MsgBox("This will keep all listed PPRs including the PPR being edited.", vbInformation + vbOKCancel, "PPR") = vbCancel Then Exit Sub
DoCmd.Close acForm, "frmdupeview"
End Sub

Private Sub btnNone_Click()
If MsgBox("This will archive all listed PPRs and keep the current PPR being edited.", vbInformation + vbOKCancel, "PPR") = vbCancel Then Exit Sub

    With rsDUPE: Do While Not .EOF
        If !PPR <> Forms!new_ppr!PPR Then
            .edit
            !archive = True
            !Status = "Cancelled"
            If !spID <> 0 Then
                qdf.Parameters("mtbid") = !spID
                Set RS = qdf.OpenRecordset
                
                If RS.RecordCount = 1 Then
                    RS.delete
                    RS.Close
                End If
            End If
            .update
        End If
        .MoveNext
    Loop: End With
    DoCmd.Close acForm, "frmdupeview"
    
End Sub

Private Sub btnSelection_Click()
If listDupe.ItemsSelected.Count = 0 Then Exit Sub
If MsgBox("This will keep selected PPRs and archive the rest.", vbInformation + vbOKCancel, "PPR") = vbCancel Then Exit Sub

    With rsDUPE: Do While Not .EOF 'iterate through complete dupe list
        match = False
        With listDupe: For Each i In .ItemsSelected 'iterate through list of chosen dupes
        
            If rsDUPE!PPR = .Column(0, i) Then 'If particular dupe from master is selected
                match = True 'selectand move on to additional actions
                Exit For
            End If
            'Else: Move to next selected supe
        Next: End With
        
        If Not match And !PPR <> Forms!new_ppr!PPR Then 'Do the following to selected dupe, if it is not locked for editing
            .edit
            !archive = True
            !Status = "Cancelled"
            
            If !spID <> 0 Then
                qdf.Parameters("mtbid") = !spID
                Set RS = qdf.OpenRecordset
                
                If RS.RecordCount = 1 Then
                    RS.delete
                    RS.Close
                End If
            End If
            .update
        End If
        
        .MoveNext 'On to the next dupe
    Loop: End With
    Forms!new_ppr.bDELETE = True
    DoCmd.Close acForm, "frmdupeview"

End Sub

Private Sub cancel_Click()
Forms!new_ppr.dupeCheckFail = True
DoCmd.Close acForm, "frmDupeView"
End Sub

Private Sub Form_Load()
'With Forms!new_ppr
'    If IsNull(OpenArgs) Then
'        DoCmd.close acForm, "frmdupeview"
'        errHandler err, "Form was not passed recordset.", "frmDupeView"
'        Exit Sub
'    End If
'
'End With

Set rsDUPE = CurrentDb.OpenRecordset(OpenArgs)
If Not rsDUPE.EOF Then
    With rsDUPE
        .MoveLast
        .MoveFirst
        Select Case .RecordCount
        
        Case 1 And !PPR = Forms!new_ppr!PPR
            DoCmd.Close acForm, "frmdupeview"
            Exit Sub
        Case 0
            DoCmd.Close acForm, "frmdupeview"
            Exit Sub
        Case Else
            
        End Select
        listDupe.RowSource = OpenArgs
        Set qdf = CurrentDb.QueryDefs("qMissionTrackerByID")
    End With
Else
    DoCmd.Close acForm, "frmdupeview"
End If
End Sub

