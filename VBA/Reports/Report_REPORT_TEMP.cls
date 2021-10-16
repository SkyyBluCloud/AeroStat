VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Report_REPORT_TEMP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private reportDate As Variant

Private Sub Report_Open(cancel As Integer)
On Error GoTo errtrap

CurrentDb.Execute "qTrafficFlightStatusUpdate", dbFailOnError

'reportDate = Null
reportDate = Parent.Parent.reportDate

Dim newSQL As String, db As DAO.Database: Set db = CurrentDb
newSQL = Replace(db.QueryDefs("qTraffic").sql, "[varDate]", "#" & reportDate & "#")
Me.RecordSource = newSQL

'Me.RecordSource = "SELECT tblTraffic.*, tblUserAuth.opInitials FROM (tblTraffic LEFT JOIN tblShiftCertifier ON tblTraffic.certifierID = tblShiftCertifier.ID) LEFT JOIN tblUserAuth ON tblShiftCertifier.username = tblUserAuth.username WHERE (((tblTraffic.Destination)=DLookup("data","tblSettings","key = 'station'")) AND ((DateValue([DOF]+Nz([ATA],Nz([ETA],Nz([ATD],[ETD])+[ETE]))))=#" & Nz(reportDate, Date) & "#)) OR (((tblTraffic.depPoint)=DLookup("data","tblSettings","key = 'station'")) AND ((DateValue([DOF]+Nz([ATD],[ETD])))=#" & Nz(reportDate, Date) & "#));"
'Me.RecordSource = "SELECT tblTraffic.*, tblShiftCertifier.opInitials FROM (tblShiftCertifier RIGHT JOIN tblUserAuth ON tblShiftCertifier.username = tblUserAuth.username) INNER JOIN tblTraffic ON tblUserAuth.opInitials = tblTraffic.AMOPS WHERE (((tblTraffic.Destination)=DLookup("data","tblSettings","key = 'station'")) AND ((DateValue([DOF])+Nz([ATA],Nz([ETA],Nz([ATD],[ETD])+[ETE])))=#" & Nz(rd, Date) & "#)) OR (((tblTraffic.depPoint)=DLookup("data","tblSettings","key = 'station'")) AND ((DateValue([DOF]+Nz([ATD],[ETD])))=#" & Nz(rd, Date) & "#));"
'Me.RecordSource = "SELECT tblTraffic.*, tblShiftCertifier.reviewSigID FROM tblShiftCertifier RIGHT JOIN tblTraffic ON tblShiftCertifier.reviewSigID = tblTraffic.reviewSigID WHERE (((tblTraffic.Destination)=DLookup("data","tblSettings","key = 'station'")) AND ((DateValue([DOF])+ Nz([ATA],Nz([ETA],Nz([ATD],[ETD])+[ETE])))=#" & Nz(rd, Date) & "#)) OR (((tblTraffic.depPoint)=DLookup("data","tblSettings","key = 'station'")) AND ((DateValue([DOF]+Nz([ATD],[ETD])))=#" & Nz(rd, Date) & "#));"
'Me.RecordSource = "SELECT tblTraffic.* FROM tblTraffic WHERE DateSerial(Year(nz(tblTraffic.depDate,Date())), Month(nz(tblTraffic.depDate,Date())), Day(nz(tblTraffic.depDate,Date()))) = #" & Nz(rd, Date) & "#"
'=Format(DateValue([dof])+Nz([ATA],Nz([ETA],Nz([ATD],[ETD])+[ETE])),"dd mmm yyyy")
sexit:
    Exit Sub
errtrap:
    ErrHandler err, Error$, "rTraffic.report_Open"
    Resume Next
End Sub




