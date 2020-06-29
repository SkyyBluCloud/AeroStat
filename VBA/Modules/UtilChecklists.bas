Attribute VB_Name = "UtilChecklists"
Option Compare Database
Const mName As String = "UtilChecklists"

Public Function isComplete(ByVal checklist As Integer, ByVal instance As Integer) As Boolean
Dim rCount, cCount As Integer
Dim criteria As String
criteria = "checklistID = " & checklist & " AND instance = " & instance
rCount = DCount("ID", "tblChecklistCompletionData", criteria)
cCount = DCount("opinitials", "tblChecklistCompletionData", criteria)
isComplete = (rCount <> 0) And (rCount = cCount)
End Function

Public Function isclosed(ByVal checklistID As Integer, ByVal instance As Integer) As Boolean
isclosed = Not IsNull(DLookup("opSig", "tblChecklistCompletionData", "checklistID = " & checklistID & " AND instance = " & instance))
End Function

Public Function startChecklist(ByVal checklistID As Integer, ByVal shiftID As Integer) As Integer
'Initiate the <checklistID> for <shiftID>
'Returns: True if successful, False if not
On Error GoTo errtrap
Dim rsItems As DAO.Recordset
Dim rsCD As DAO.Recordset
Dim instance As Integer
If IsNull(DLookup("shiftid", "tblshiftmanager", "shiftid = " & shiftID)) Then GoTo fExit
instance = Nz(DMax("instance", "tblChecklistCompletionData", "checklistID = " & checklistID), 0) + 1
Set rsItems = CurrentDb.OpenRecordset("SELECT * FROM tblCheckListItems WHERE checklistID = " & checklistID & " ORDER BY tblChecklistItems.order")
Set rsCD = CurrentDb.OpenRecordset("tblChecklistCompletionData")

    With rsItems: Do While Not .EOF
        With rsCD
            .AddNew
            !instance = instance
            !checklistID = checklistID
            !itemID = rsItems!itemID
            !shiftID = shiftID
            !startDate = Now
            .Update
        End With
        .MoveNext
    Loop: End With
    
    startChecklist = instance
    
fExit:
    log "Started instance " & instance & " of checklist " & checklistID & " for shift " & shiftID, "UtilChecklists.startChecklist"
    Exit Function
    
errtrap:
    ErrHandler err, Error$, mName & ".startChecklist"
    
End Function

Public Function closeChecklist(ByVal checklistID As Integer, ByVal instance As Integer, Optional ByVal opInitials As String) As Boolean
If IsNull(instance) Then instance = DMax("instance", "tblChecklistCompletionData", "checklistID = " & checklistID)
If IsNull(opInitials) Then opInitials = Util.getOpInitials
Dim db As DAO.Database
Set db = CurrentDb

    opSig = Util.getUSN(opInitials)
    db.Execute "UPDATE tblChecklistCompletionData SET opSig = '" & opSig & "' WHERE checklistID = " & checklistID & " AND instance = " & instance
    closeChecklist = db.RecordsAffected <> 0
    
End Function

Public Function deleteChecklist(ByVal checklistID As Integer, ByVal shiftID As Integer, Optional ByVal instance As Integer) As Boolean
'Delete [instance|last instance] version of <checklistID> within <shiftID>
'Returns: True if successful, False if not
'Dim rs As DAO.Recordset
'Set rs = CurrentDb.OpenRecordset("SELECT * FROM tblChecklistCompletionData WHERE checklistID = " & checklistID & " AND shiftID = " & shiftID & " AND instance = " & instance)

'    With rs
'        If .EOF Then Exit Function
'        Do While Not .EOF
'            .delete
'            .MoveNext
'        Loop
'    End With
On Error GoTo errtrap
Dim db As DAO.Database
Set db = CurrentDb

    If instance = 0 Then instance = Nz(DMax("instance", "tblChecklistCompletionData", "checklistID = " & checklistID), 0)
    If instance = 0 Then GoTo fExit

    db.Execute ("DELETE * FROM tblChecklistCompletionData WHERE checklistID = " & checklistID & " AND shiftID = " & shiftID & " AND instance = " & instance)
    deleteChecklist = db.RecordsAffected <> 0
    
fExit:
    log CStr(deleteChecklist), "UtilChecklists.startChecklist"
    Exit Function

errtrap:
    ErrHandler err, Error$, mName & ".deleteChecklist"
    GoTo fExit
End Function

