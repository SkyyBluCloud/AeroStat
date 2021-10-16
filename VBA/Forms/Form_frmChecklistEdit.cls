VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmChecklistEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Public bClose As Boolean
Private currChecklist As Integer

Private Sub btnClear_Click()
    Me.filter = "itemID = " & 0
    Me.FilterOn = True
End Sub

Private Sub btnDelete_Click()
On Error Resume Next
    Dim db As DAO.Database
    Set db = CurrentDb
    If IsNull(lstItems) Then Exit Sub

    If MsgBox("Delete item?", vbQuestion + vbYesNo, Me.Caption) = vbYes Then
        bClose = True
        db.Execute "DELETE FROM tblChecklistItems WHERE itemID = " & lstItems
        Util.log "Deleted " & db.RecordsAffected & " item(s) from the checklist index.", Me.Name & ".btnDelete_Click"
        'DoCmd.RunCommand acCmdDeleteRecord
    End If
    
    Me.Requery
    lstItems.Requery
End Sub

Private Sub btnMode_AfterUpdate()
If NewRecord Then Exit Sub
currChecklist = checklistID
    Select Case btnMode
    Case True 'Edit mode; clear anything type in and not saved
        Undo
        btnMode.Caption = "Edit"
    Case False 'Add mode; move to new record
        Undo
        btnMode.Caption = "Add"
        DoCmd.GoToRecord , "", acNewRec
        checklistID = currChecklist
        order = DMax("order", "tblchecklistitems", "checklistid = " & checklistID) + 1
    End Select
End Sub

Private Sub btnMode_Click()
'If NewRecord Then Exit Sub
'currChecklist = checklistID
'    Select Case btnMode
'    Case True 'Edit mode; clear anything type in and not saved
'        Undo
'        btnMode.Caption = "Edit"
'    Case False 'Add mode; move to new record
'        Undo
'        btnMode.Caption = "Add"
'        DoCmd.GoToRecord , "", acNewRec
'        checklistID = currChecklist
'        order = DMax("order", "tblchecklistitems", "checklistid = " & checklistID) + 1
'    End Select
End Sub

Private Sub btnSaveItem_Click()
On Error Resume Next
bClose = True
If Dirty Then Dirty = False
Parent.refreshChecklistEditor
End Sub

Private Sub Form_BeforeUpdate(cancel As Integer)
    If Not bClose Then
        cancel = True
        Undo
        Exit Sub
    End If
End Sub

Private Sub Form_Current()
    btnMode = Not NewRecord
    If Not NewRecord Then btnMode_AfterUpdate
    bClose = False
End Sub

Private Sub lstItems_AfterUpdate()
'    If NewRecord Then DoCmd.SearchForRecord , "", , "itemID = " & lstItems
'    DoCmd.SearchForRecord , "", , "itemID = " & lstItems
    Me.filter = "itemID = " & lstItems
    Me.FilterOn = True
End Sub

'Public Function refreshChecklistEditor()
'On Error GoTo errtrap
'
'    If Nz(currChecklist) = "" Then
'        'Me.sfEditChecklist.SourceObject = ""
'        Exit Function
'    Else
'        Dim qdf As DAO.QueryDef
'        Set qdf = CurrentDb.QueryDefs("qChecklistItems")
'        qdf.Parameters("checklist") = currChecklist
'
'        'Me.sfEditChecklist.SourceObject = "frmChecklistEdit"
'        With Me
'            .filter = "checklistID = " & currChecklist
'            .FilterOn = True
'            Set lstItems.Recordset = qdf.OpenRecordset
'        End With
'    End If
'
'sExit:
'    Exit Function
'errtrap:
'    errHandler err, Error$, Me.Name & ".refreshChecklistEditor"
'End Function
