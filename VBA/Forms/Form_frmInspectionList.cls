VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmInspectionList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Dim shiftID As Integer

Private Sub Form_Load()
    shiftID = Nz(Parent.Parent.shiftID, 0)
    populate shiftID
End Sub

Private Sub inspList_DblClick(cancel As Integer)

    Parent.changeView "frmInspectionNew", shiftID, inspList.Column(0)
    'inspList.Requery
End Sub

Private Sub newCheckInsp_Click()
    'DoCmd.OpenForm "frmInspectionNew", , , , acFormAdd, acDialog

    Parent.changeView "frmInspectionNew", shiftID
    
End Sub

Private Sub timeBlock_Click()
Dim qDate As String
    Select Case timeBlock
        Case 1
            populate
            Exit Sub
        Case 2
            populate shiftID
            Exit Sub
        Case 3
            date1 = Format("1/1/" & Year(Now), "dd-mmm-yy")
            date2 = Format("3/31/" & Year(Now), "dd-mmm-yy")
        Case 4
            date1 = Format("4/1/" & Year(Now), "dd-mmm-yy")
            date2 = Format("6/30/" & Year(Now), "dd-mmm-yy")
        Case 5
            date1 = Format("7/1/" & Year(Now), "dd-mmm-yy")
            date2 = Format("9/30/" & Year(Now), "dd-mmm-yy")
        Case 6
            date1 = Format("10/1/" & Year(Now), "dd-mmm-yy")
            date2 = Format("12/31/" & Year(Now), "dd-mmm-yy")
    End Select
    qDate = "Between Format('" & date1 & "','mmddyy') And Format('" & date2 & "','mmddyy')"
    populate , qDate
End Sub

Private Function populate(Optional ByVal cond As String, Optional scope As String)
Dim q As String
q = "SELECT tblInspections.ID, tblInspections.shiftID, tblInspections.Inspection, tblInspections.checkDate, tblInspections.outTime, tblInspections.inTime, tblInspections.callsign, tblInspections.initials, tblInspections.Remarks, tblInspections.AF483Count FROM tblInspections"

    If Nz(scope) <> "" Then
        q = q & " WHERE Format([checkDate],'mmddyy') " & scope
    ElseIf Nz(cond) <> "" Then
        q = q & " WHERE shiftID = " & cond
    End If
    inspList.RowSource = q
End Function
