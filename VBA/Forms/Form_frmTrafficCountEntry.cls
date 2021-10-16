VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmTrafficCountEntry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
Dim bClose As Boolean

Private Sub newCount(Optional ByVal varDate As Variant)
Dim rsItems As DAO.Recordset: Set rsItems = CurrentDb.OpenRecordset("tblAnalyticsItems")
Dim rsData As DAO.Recordset: Set rsData = CurrentDb.OpenRecordset("tblAnalyticsData")
log "New count!", Me.Name & ".newCount"

    With rsItems: Do While Not .EOF
        With rsData
            .AddNew
            !itemID = rsItems!ID
            !countDate = varDate
            .update
        End With
        .MoveNext
    Loop: End With
    rsData.Close
    Set rsData = Nothing
    rsItems.Close
    Set rsItems = Nothing
    
    lstItems.Requery
    'Call lstItems_AfterUpdate
End Sub

Private Sub btnDelete_Click()
If MsgBox("Delete count for this date?", vbQuestion + vbYesNo, "New Count") = vbYes Then
    Me.Visible = False
    CurrentDb.Execute "DELETE FROM tblAnalyticsData WHERE countDate = #" & ctlCountDate & "#"
    log "Deleted traffic count data for " & ctlCountDate, Me.Name & ".btnDelete_Click", "WARN"
    DoCmd.Close acForm, Me.Name
End If
End Sub

Private Sub btnDone_Click()
If Dirty Then Dirty = False
DoCmd.Close acForm, Me.Name
End Sub

Private Sub Form_Current()
bClose = False
End Sub

Private Sub Form_Load()
'Check if count for OpenArgs(or today) exists
Dim cd As Variant: cd = DLookup("countDate", "tblAnalyticsData", "countDate = #" & Nz(OpenArgs, Date) & "#")
If IsNull(cd) Then
    newCount Nz(OpenArgs, Date)
Else
    ctlCountDate = cd
End If
End Sub

Private Sub itemCount_KeyDown(KeyCode As Integer, Shift As Integer)

    Select Case KeyCode
    Case vbKeyReturn, vbKeyDown
        bClose = True
        If Dirty Then Dirty = False
        If lstItems < lstItems.ListCount - 1 Then
            lstItems = lstItems + 1
        Else
            lstItems = 0
        End If
        lstItems.Requery
        Call lstItems_AfterUpdate
        itemCount.SetFocus
    
    Case vbKeyUp
        bClose = True
        If Dirty Then Dirty = False
        
        If lstItems <> 0 Then
            lstItems = lstItems - 1
        End If
        lstItems.Requery
        Call lstItems_AfterUpdate
        itemCount.SetFocus
    
     End Select
        
        
End Sub

Private Sub lstItems_AfterUpdate()
    
    Me.filter = "ID = " & lstItems.Column(0)
    Me.FilterOn = True
    
    
End Sub

Private Sub lstItems_BeforeUpdate(cancel As Integer)
If Not bClose Then
    'Cancel = True
    Undo
    Exit Sub
End If
End Sub
