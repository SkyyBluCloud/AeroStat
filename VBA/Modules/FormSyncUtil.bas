Attribute VB_Name = "FormSyncUtil"
Option Compare Database
Option Explicit

Private Sub init()
'First time table setup.
On Error GoTo errtrap
Dim db As DAO.Database: Set db = CurrentDb
Dim c As Container: Set c = db.Containers("Forms")

    Util.trunc "tblFormSyncGlobal", True
    Util.trunc "tblFormSyncLocal", True
    
    log "Updating records...", "FormSyncUtil.init"

    Dim frm: For Each frm In c.Documents
        If Left(frm.Name, 1) <> "@" Then
            db.Execute "INSERT INTO tblFormSyncGlobal(formName) " & _
                                "SELECT '" & frm.Name & "'", dbFailOnError
            db.Execute "INSERT INTO tblFormSyncLocal(formName) " & _
                                "SELECT '" & frm.Name & "'", dbFailOnError
            
        End If
    Next frm
    log "Done!", "FormSyncUtil.init"
    
sexit:
    Exit Sub
errtrap:
    ErrHandler err, Error$, "FormSyncUtil.init"
End Sub

Public Sub syncForm(ByVal frm As String, Optional ByVal clientOnly As Boolean)
'Synchronizes the specified form.
'clientOnly: True = Update local table ONLY; False = Update both global and local tables (and subsequently trigger everyone else to update)
On Error GoTo errtrap

Dim n As Date: n = Now
Dim db As DAO.Database: Set db = CurrentDb
'If isFormSynced(frm) And clientOnly Then
'    Exit Sub
'End If

    If Not clientOnly Then
        db.Execute "UPDATE tblFormSyncGlobal SET syncTime = #" & n & "# WHERE formName = '" & frm & "'", dbFailOnError
    End If
    
    db.Execute "UPDATE tblFormSyncLocal SET syncTime = #" & _
        IIf(clientOnly, _
            DLookup("syncTime", "tblFormSyncGlobal", "formName = '" & frm & "'"), _
            n) & _
        "# WHERE formName = '" & frm & "'", dbFailOnError
        
    log IIf(clientOnly, "Received update from '", "Pushed global update for '") & frm & "'", "FormSyncUtil.syncForm", "UPDATE"
    
sexit:
    Set db = Nothing
    Exit Sub
errtrap:
    ErrHandler err, Error$, "FormSyncUtil.syncForm"
    Resume sexit
End Sub

Public Function isFormSynced(ByVal frm As String) As Boolean
'Checks if a form is in sync.
isFormSynced = _
DLookup("syncTime", "tblFormSyncLocal", "formName = '" & frm & "'") = DLookup("syncTime", "tblFormSyncGlobal", "formName = '" & frm & "'")
End Function
