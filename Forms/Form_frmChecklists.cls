VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmChecklists"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Private currentChecklist As Integer
Private currentInstance As Integer

Private Sub btnCheck_Click()
If Me.lstDetails = "" Then Me.lstDetails = 1
Dim suppInitials As String
Dim ccdID, itemID, checklistID, instance As Integer
Dim si As Variant
Dim siArray() As String
Dim exe As String
ccdID = Me.lstDetails.Column(0)
itemID = Me.lstDetails.Column(1)
checklistID = currentChecklist
instance = currentInstance
si = DLookup("suppInitials", "tblchecklistitems", "itemid = " & itemID)
    
    'Check if shift is closed and signed; if true, restrict to amos
    If DLookup("closed", "tblShiftManager", "shiftid = " & shiftID) Then
        If Util.getOpInitials <> DLookup("right(superlead,2)", "tblshiftmanager", "shiftid = " & shiftID) Then
            MsgBox "This shift is closed. Only the AMOS can make changes.", vbInformation, "Checklist"
            Exit Sub
        End If
    End If
    
    'Check if already accomplished
    If Not IsNull(DLookup("completeDate", "tblchecklistcompletiondata", "ID = " & ccdID)) Then
        If MsgBox("This was already completed. Re-accomplish?", vbInformation + vbYesNo, "Checklist") = vbNo Then Exit Sub
    End If

    'Handle Coordination if necessary
    If Not IsNull(si) Then
        siArray = Split(si, ",")
        'suppInitials = InputBox("Coordination is required. Enter the initials below:" & vbCrLf & si, "Coordination", join(siArray, "/__, ") & "/__")
        
        For Each s In siArray
            suppInitials = Nz(suppInitials) & IIf(IsNull(suppInitials), "", ", ") & InputBox("Call/notify and enter initials for " & s, "Coordination Required")
        Next
    End If
    
    '???
    'exe = "UPDATE tblChecklistCompletionData " & _
                "SET opInitials = '" & Util.getOpInitials & "', " & _
                "completeDate = format('" & Now & "','dd-mmm-yy hh:nn')" & _
                IIf(Not IsNull(si), ", suppInitials = '" & suppInitials & "'", "") & _
            " WHERE ID = " & ccdID
                        
    'Profit
    
    'CurrentDb.Execute exe
    opInitials = Nz(opInitials, Util.getOpInitials)
    completeDate = Format(Now, "dd-mmm-yy hh:nn")
                        
    If Dirty Then Dirty = False
    
    'On to the next item
    Me.lstDetails = Me.lstDetails + 1
    lstDetails_AfterUpdate
    lstDetails.Requery
End Sub

Private Sub btnClear_Click()
If IsNull(lstDetails) Then Exit Sub
If MsgBox("Clear entry?", vbQuestion + vbYesNo, "Checklist") = vbNo Then Exit Sub

    completeDate = ""
    opInitials = ""
    Remarks = ""
    If Dirty Then Dirty = False
    lstDetails.Requery
End Sub

Private Sub btnCloseCklist_Click()
If IsNull(currentChecklist) Then Exit Sub

    If Not UtilChecklists.isComplete(lstOpenChecklists.Column(1), currentChecklist) Then
        MsgBox "Can not close incomplete checklist.", vbCritical, "Checklist"
        Exit Sub
        
    Else
        If MsgBox("This will close and sign the current checklist. Continue?", vbQuestion + vbYesNo, "Checklist") = vbYes Then
            If UtilChecklists.closeChecklist(lstOpenChecklists.Column(1), currentChecklist) Then
                refreshLists
            End If
        End If
    End If
    lstOpenChecklists.Requery
End Sub

Private Sub btnDeleteCklist_Click()
If Me.lstOpenChecklists = "" Then Exit Sub
Dim inst As Integer
Dim cid As Integer
inst = Me.lstOpenChecklists
cid = Me.lstOpenChecklists.Column(1)

    
    If Not IsNull(DLookup("opsig", "tblchecklistcompletiondata", "instance = " & Me.lstOpenChecklists & " AND checklistID = " & Me.lstOpenChecklists.Column(1))) Then
        If DLookup("opsig", "tblchecklistcompletiondata", "instance = " & Me.lstOpenChecklists & " AND checklistID = " & Me.lstOpenChecklists.Column(1)) <> Util.getUSN Then
            MsgBox "This checklist was already signed. Remove the signature and try again.", vbInformation, "Shift Board"
            Exit Sub
        End If
    ElseIf MsgBox("Delete?", vbQuestion + vbYesNo, "Shift Board") = vbNo Then
        Exit Sub
    End If

    UtilChecklists.deleteChecklist cid, shiftID, inst
    FilterOn = False
    refreshLists
    For Each ctl In Controls
        If ctl.Tag = "itemselected" Then ctl.Visible = (Nz(Me.detailView) <> "")
    Next
    currentChecklist = 0
End Sub

Private Sub btnStartCklist_Click()
If Me.lstAvailableChecklists = "" Then Exit Sub
Dim inst As Integer

    inst = UtilChecklists.startChecklist(Me.lstAvailableChecklists, shiftID)
    Me.lstOpenChecklists = inst
    refreshLists
    lstOpenChecklists_AfterUpdate
    Me.lstAvailableChecklists = ""
    'DoCmd.SearchForRecord , "", acFirst, "checklistID = " & currentChecklist & " AND instance = " & inst

End Sub

Private Sub Form_Load()
shiftID = Nz(Me.OpenArgs, 0)
If Parent.Name = "frmShiftBoard" Then shiftID = Nz(Parent.shiftID, Parent.OpenArgs)
shiftID = Nz(shiftID, 0)
populate
End Sub

Private Sub lstDetails_AfterUpdate()
'DoCmd.SearchForRecord , "", acFirst, "ID = " & lstDetails.Column(0)
Me.filter = "ID = " & Nz(lstDetails.Column(0), 0)
Me.FilterOn = True
'Me.Remarks = Me.lstDetails.Column(8)
End Sub

Private Sub lstDetails_Click()
For Each ctl In Controls
    If ctl.Tag = "itemselected" Then ctl.Visible = (Me.detailView <> "")
Next
End Sub

Private Sub lstOpenChecklists_AfterUpdate()
currentChecklist = lstOpenChecklists
currentInstance = lstOpenChecklists
Me.lblDeatils.Caption = Me.lstOpenChecklists.Column(3)
lstDetails.RowSource = "SELECT tblChecklistCompletionData.ID, tblChecklistItems.itemID, tblChecklistCompletionData.instance, tblChecklistCompletionData.checklistID, tblChecklistItems.order, tblChecklistCompletionData.startDate, tblChecklistCompletionData.completeDate, tblChecklistItems.item, tblChecklistCompletionData.remarks, tblChecklistCompletionData.opInitials " & _
                        "FROM tblChecklistItems INNER JOIN tblChecklistCompletionData ON tblChecklistItems.itemID = tblChecklistCompletionData.itemID " & _
                        "WHERE (((tblChecklistCompletionData.instance)=" & Nz(lstOpenChecklists, 0) & ")) AND (((tblChecklistCompletionData.checklistID)=" & Nz(lstOpenChecklists.Column(1), 0) & ")) ORDER BY tblChecklistItems.order;"
lstDetails_AfterUpdate
End Sub

Private Sub refreshLists()
For Each ctl In Controls
    If TypeOf ctl Is ListBox Then ctl.Requery
Next
End Sub

Private Sub populate()
lstOpenChecklists.RowSource = "SELECT tblChecklistCompletionData.instance, tblChecklistIndex.checklistID, tblChecklistIndex.code, tblChecklistIndex.title, IIf([opsig] Is Not Null,'SIGNED','') AS Signed " & _
                                "FROM tblChecklistIndex INNER JOIN tblChecklistCompletionData ON tblChecklistIndex.checklistID = tblChecklistCompletionData.checklistID " & _
                                "WHERE (((tblChecklistCompletionData.shiftID)=" & shiftID & ")) " & _
                                "GROUP BY tblChecklistCompletionData.instance, tblChecklistIndex.checklistID, tblChecklistIndex.code, tblChecklistIndex.title, IIf([opsig] Is Not Null,'SIGNED','');"
cboInitials.RowSource = "SELECT opInitials FROM tblUserAuth WHERE lastShift = " & shiftID
End Sub

Private Sub rmkSave_Click()
If Dirty Then Dirty = False
End Sub

