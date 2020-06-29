VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmNOTAMview"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub cancel_Click()
Dim RS As DAO.Recordset
Dim NOTAM As String
Dim s() As String
Dim c As String
Dim idx As Integer
Dim selection As Boolean

For Each ctl In Controls
    If TypeOf ctl Is ListBox Then
        For i = 0 To ctl.ListCount - 1
            If ctl.Selected(i) = True Then
                selection = True
                c = ctl.Name
                Exit For
            End If
        Next
        If selection Then Exit For
    End If
Next
If Not selection Then Exit Sub
If DLookup("ntype", "tblNOTAM", "NOTAM = '" & NOTAM & "'") = "C" Then Exit Sub

    With Controls(c)
    For Each i In .ItemsSelected
        ReDim Preserve s(0 To idx) As String
        s(idx) = .Column(1, i)
        idx = idx + 1
    Next
    End With
    
    If MsgBox("Cancel NOTAM?: " & vbCrLf & join(s, ", "), vbQuestion + vbYesNo, "NOTAM Control") = vbNo Then Exit Sub

    Set RS = CurrentDb.OpenRecordset("SELECT * FROM tblNOTAM WHERE tblNOTAM.NOTAM = '" & join(s, "' OR tblNOTAM.NOTAM = '") & "'")
    With RS: Do While Not .EOF
            .edit
            !isCancelled = True
            !endTime = Now
            .Update
            .MoveNext
    Loop: End With

For Each ctl In Controls
    If TypeOf ctl Is ListBox Then ctl.Requery
Next
End Sub

Private Sub Detail_Click()
For Each ctl In Controls
    If TypeOf ctl Is ListBox Then ctl.Value = ""
Next
End Sub

Private Sub lbControl_Click()
For Each ctl In Controls
    If TypeOf ctl Is ListBox And ctl.Name <> "lbControl" Then ctl.Value = ""
Next
End Sub

Private Sub lbControl_DblClick(cancel As Integer)
DoCmd.OpenForm "frmNOTAMInput", , , "ID = " & lbControl.Column(0), acFormEdit, acDialog
lbControl.Requery
End Sub

Private Sub lbExpired_Click()
For Each ctl In Controls
    If TypeOf ctl Is ListBox And ctl.Name <> "lbExpired" Then ctl.Value = ""
Next
End Sub

Private Sub lbExpired_DblClick(cancel As Integer)
DoCmd.OpenForm "frmNOTAMInput", , , "ID = " & lbExpired.Column(0), acFormEdit, acDialog
lbExpired.Requery
End Sub

Private Sub lbLocal_Click()
For Each ctl In Controls
    If TypeOf ctl Is ListBox And ctl.Name <> "lbLocal" Then ctl.Value = ""
Next
End Sub

Private Sub lbLocal_DblClick(cancel As Integer)
DoCmd.OpenForm "frmNOTAMInput", , , "ID = " & lbLocal.Column(0), acFormEdit, acDialog
lbLocal.Requery
End Sub

Private Sub lbSafety_Click()
For Each ctl In Controls
    If TypeOf ctl Is ListBox And ctl.Name <> "lbSafety" Then ctl.Value = ""
Next
End Sub

Private Sub lbSafety_DblClick(cancel As Integer)
DoCmd.OpenForm "frmNOTAMInput", , , "ID = " & lbSafety.Column(0), acFormEdit, acDialog
lbSafety.Requery
End Sub

Private Sub newNotam_Click()
DoCmd.OpenForm "frmNOTAMParse", , , , , acDialog
For Each ctl In Controls
    If TypeOf ctl Is ListBox Then ctl.Requery
Next
End Sub

Private Sub ogSeries_Click()
Dim where As String

    Select Case ogSeries
    Case 1 'M
        where = "WHERE left(tblNOTAM.NOTAM,1) = 'M'"
        
    Case 2 'L
        where = "WHERE left(tblNOTAM.NOTAM,1) = 'L'"
        
    Case 3 'V
        where = "WHERE left(tblNOTAM.NOTAM,1) = 'V'"
        
    Case 4 'LOA (Issued NOTAM for backup facility)
        where = "WHERE tblNOTAM.aerodrome <> '" & DLookup("station", "tblSettings") & "'"
        
    End Select
    
    lbControl.RowSource = "SELECT tblNOTAM.ID, tblNOTAM.NOTAM, tblNOTAM.qcode, tblNOTAM.startTime, tblNOTAM.endTime, tblNOTAM.period, tblNOTAM.verbiage, tblNOTAM.issuedBy FROM tblNOTAM " & where & " ORDER BY Right(tblNOTAM.NOTAM,2) DESC , tblNOTAM.NOTAM DESC , tblNOTAM.startTime DESC , tblNOTAM.endTime DESC;"
    
End Sub
