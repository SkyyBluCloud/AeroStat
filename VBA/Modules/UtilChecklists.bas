Attribute VB_Name = "UtilChecklists"
Option Compare Database
Const mName As String = "UtilChecklists"

Public Function isComplete(ByVal instance As Integer, ByVal checklistID As Integer) As Boolean
Dim rCount, cCount As Integer
Dim criteria As String

    criteria = "instance = " & instance
    rCount = DCount("ID", "tblChecklistItemsData", criteria)
    cCount = DCount("opInitials", "tblChecklistItemsData", criteria & " AND nz(opInitials) <> ''")
    
    isComplete = (rCount <> 0) And (rCount = cCount)
End Function

Public Function isClosed(ByVal checklistID As Integer, ByVal instance As Integer) As Variant
isClosed = DLookup("certifierID", "qShiftChecklists", "checklistID = " & checklistID & " AND instance = " & instance) <> 0
End Function

Public Function startChecklist(ByVal checklistID As Integer, ByVal shiftID As Integer) As Integer
'Initiate the <checklistID> for <shiftID>
'Returns the instance of newly started checklist; returns 0 if unsuccessful
On Error GoTo errtrap
Dim rsItems As DAO.Recordset
Dim rsCD As DAO.Recordset

If IsNull(DLookup("shiftid", "tblshiftmanager", "shiftid = " & shiftID)) Then GoTo fexit

Dim instance As Integer
instance = Nz(CurrentDb.OpenRecordset("SELECT Max([instance]) FROM qShiftChecklists WHERE checklistID = " & checklistID).Fields(0), 0) + 1

'Set rsItems = CurrentDb.OpenRecordset("SELECT * FROM tblCheckListItems WHERE checklistID = " & checklistID & " ORDER BY tblChecklistItems.order")
'Set rsCD = CurrentDb.OpenRecordset("tblChecklistItemsData")

    Dim db As DAO.Database: Set db = CurrentDb
    db.Execute _
    "INSERT INTO tblChecklistItemsData (instance, itemID, shiftID, startDate) " & _
    "SELECT " & instance & ", tblChecklistItems.itemID, " & shiftID & ", Now() " & _
    "FROM tblChecklistItems WHERE checklistID = " & checklistID & " ORDER BY tblChecklistItems.order", dbFailOnError

'    With rsItems: Do While Not .EOF
'        With rsCD
'            .AddNew
'            !instance = instance
'            !itemID = rsItems!itemID
'            !shiftID = shiftID
'            !startDate = Now
'            .update
'        End With
'        .MoveNext
'    Loop: End With
    
    startChecklist = instance
    
fexit:
    'rsItems.Close
    'Set rsItems = Nothing
    
    'rsCD.Close
    'Set rsCD = Nothing
    Set db = Nothing
    
    log "Started instance " & instance & " of checklist " & checklistID & " for shift " & shiftID, "UtilChecklists.startChecklist"
    Exit Function
    
errtrap:
    ErrHandler err, Error$, mName & ".startChecklist"
    
End Function

Public Function closeChecklist(ByVal checklistID As Integer, ByVal instance As Integer, Optional ByVal opInitials As String) As Boolean
On Error GoTo fErr
If IsNull(instance) Then Exit Function
If IsNull(opInitials) Then opInitials = Util.getOpInitials(getUSN)
Dim db As DAO.Database: Set db = CurrentDb
    
    Dim cert As Variant: cert = UtilCertifier.newCert(getUSN)
    db.Execute "UPDATE tblChecklistItemsData INNER JOIN tblChecklistItems ON tblChecklistItemsData.itemID = tblChecklistItems.itemID " & _
                "SET certifierID = " & cert & " WHERE checklistID = " & checklistID & " AND instance = " & instance & ";", dbFailOnError
                
    closeChecklist = db.RecordsAffected <> 0
    
fexit:
    Set db = Nothing
    Exit Function
fErr:
    ErrHandler err, Error$, "UtilChecklists.closeChecklist"
    
End Function

Public Function deleteChecklist(ByVal checklistID As Integer, ByVal shiftID As Integer, Optional ByVal instance As Integer) As Boolean
'Delete [instance|last instance] version of <checklistID> within <shiftID>
'Returns: True if successful, False if not
'Dim rs As DAO.Recordset
'Set rs = CurrentDb.OpenRecordset("SELECT * FROM tblChecklistItemsData WHERE checklistID = " & checklistID & " AND shiftID = " & shiftID & " AND instance = " & instance)

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

    If instance = 0 Then instance = Nz(DMax("instance", "tblChecklistItemsData", "checklistID = " & checklistID), 0)
    If instance = 0 Then GoTo fexit

    db.Execute ("DELETE * FROM tblChecklistItemsData WHERE shiftID = " & shiftID & " AND instance = " & instance)
    deleteChecklist = db.RecordsAffected <> 0
    
fexit:
    'log CStr(deleteChecklist), "UtilChecklists.startChecklist"
    Exit Function

errtrap:
    ErrHandler err, Error$, mName & ".deleteChecklist"
    GoTo fexit
End Function

