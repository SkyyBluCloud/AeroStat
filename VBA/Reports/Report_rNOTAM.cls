VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Report_rNOTAM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
Private reportDate As Variant

Private Sub btnPrint_Click()
DoCmd.RunCommand acCmdPrint
End Sub

Private Sub Report_Open(cancel As Integer)
On Error GoTo errtrap
reportDate = Parent.Parent.reportDate
Me.RecordSource = "SELECT tblNOTAM.* FROM tblNOTAM WHERE datevalue(tblNOTAM.startTime) = #" & Nz(reportDate, Date) & "#;"

    Dim newSQL As String, db As DAO.Database: Set db = CurrentDb
    newSQL = Replace(db.QueryDefs("qNOTAMReport").sql, "[varDate]", "#" & reportDate & "#")
    Me.RecordSource = newSQL
sexit:
    Exit Sub
errtrap:
    Select Case err
        Case 2452: Resume Next 'No-Parent error
        
    End Select
    
    ErrHandler err, Error$, "rNOTAM.Report_Open"
    Resume Next
End Sub


