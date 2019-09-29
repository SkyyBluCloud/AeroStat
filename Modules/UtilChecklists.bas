Attribute VB_Name = "UtilChecklists"
Option Compare Database
Const mName As String = "UtilChecklists"

Public Function startChecklist(ByVal checklistID As Integer, ByVal shiftID As Integer) As Integer
'Initiate the <checklistID> for <shiftID>
'Returns: True if successful, False if not
On Error GoTo errTrap
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
            .update
        End With
        .MoveNext
    Loop: End With
    
    startChecklist = instance
    
fExit:
    log CStr(startChecklist), "UtilChecklists.startChecklist"
    Exit Function
    
errTrap:
    errHandler err, Error$, mName & ".startChecklist"
    
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
On Error GoTo errTrap
Dim db As DAO.Database
Set db = CurrentDb

    If instance = 0 Then instance = Nz(DMax("instance", "tblChecklistCompletionData", "checklistID = " & checklistID), 0)
    If instance = 0 Then GoTo fExit

    db.Execute ("DELETE * FROM tblChecklistCompletionData WHERE checklistID = " & checklistID & " AND shiftID = " & shiftID & " AND instance = " & instance)
    deleteChecklist = db.RecordsAffected <> 0
    
fExit:
    log CStr(deleteChecklist), "UtilChecklists.startChecklist"
    Exit Function

errTrap:
    errHandler err, Error$, mName & ".deleteChecklist"
    GoTo fExit
End Function

