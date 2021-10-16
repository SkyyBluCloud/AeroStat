VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmChecklistsDash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub btnDelete_Click()
On Error GoTo errtrap
If IsNull(lstChecklists) Then Exit Sub
Dim cName As String: cName = lstChecklists.Column(2)
Dim db As DAO.Database: Set db = CurrentDb

    If MsgBox("Delete " & cName & " checklist?", vbQuestion + vbYesNo, "Checklists") = vbYes Then
        db.Execute "DELETE FROM tblChecklistIndex WHERE checklistID = " & lstChecklists
        log "Deleted " & db.RecordsAffected & " checklist(s)", Me.Name & ".btnDelete_Click"
    End If
    lstChecklists = ""
    lstChecklists.Requery
    
sexit:
    Exit Sub
errtrap:
    ErrHandler err, Error$, Me.Name & ".btnDelete_Click"
End Sub

Private Sub btnNew_Click()
Dim db As DAO.Database: Set db = CurrentDb

    db.Execute "INSERT INTO tblChecklistIndex (title, code) VALUES ('New Checklist', '###')"
    Me.lstChecklists.Requery
    Me.lstChecklists = db.OpenRecordset("SELECT @@IDENTITY").Fields(0) 'Recalls field '0' from previosly executed query
    
    With sfEditChecklist.Form
        .Recordset.AddNew
        !checklistID = lstChecklists
        !order = 1
        !item = "New Item"
        .Recordset.update
        
        .bClose = True
        .filter = "checklistID = " & Me.lstChecklists
        .FilterOn = True
    End With
    
    refreshChecklistEditor
    
        
End Sub

Public Function refreshChecklistEditor()
On Error GoTo errtrap

    If Nz(lstChecklists) = "" Then
        Me.sfEditChecklist.SourceObject = ""
    Else
        Dim qdf As DAO.QueryDef
        Set qdf = CurrentDb.QueryDefs("qChecklistItems")
        qdf.Parameters("checklist") = lstChecklists
        
        Me.sfEditChecklist.SourceObject = "frmChecklistEdit"
        With Me.sfEditChecklist.Form
            .filter = "tblChecklistIndex.checklistID = " & lstChecklists
            .FilterOn = True
            Set !lstItems.Recordset = qdf.OpenRecordset
        End With
    End If
    
sexit:
    Exit Function
    Resume Next
errtrap:
    Select Case err
    Case 3709 'Filter not found
        Resume Next
    Case Else
        
    End Select
    ErrHandler err, Error$, Me.Name & ".refreshChecklistEditor"
End Function



