Attribute VB_Name = "UtilAutoCount"
Option Compare Database
Option Explicit

Public Function newCount(ByVal countDate As Date)
On Error GoTo fErr
countDate = DateValue(countDate)
Dim RS As DAO.Recordset
Dim qdf As DAO.QueryDef

    If IsNull(DLookup("countDate", "@AUTOCOUNT", "countDate = #" & countDate & "#")) Then
        Set RS = CurrentDb.OpenRecordset("@AUTOCOUNT")
        RS.AddNew
        RS!countDate = countDate
    Else
        Set RS = CurrentDb.OpenRecordset("SELECT * FROM [@AUTOCOUNT] WHERE countDate = #" & countDate & "#")
        RS.edit
    End If
    
    With RS
        'Flight Plans
        Set qdf = CurrentDb.QueryDefs("qFPtype")
        qdf.Parameters("varFPType") = 2 'DD1801s
        !countFlightPlans = qdf.OpenRecordset.Fields(0)
        
        'Stereos
        Set qdf = CurrentDb.QueryDefs("qStereo")
        qdf.Parameters("varDate") = countDate
        !countStereo = qdf.OpenRecordset.Fields(0)
        
        'NOTAM
        !countnotams = DCount("NOTAM", "tblNOTAM", "(ntype = 'N' or ntype = 'R') AND datevalue(starttime) = #" & countDate & "#")
        
        'Inspections
        '(Needs a table)
        
        'BASH
        '(Has a table, but needs work)
        
        'IFE
        '(Needs the inspections table)
        
        'PPRs
        !countPPRs = DCount("PPR", "tblPPR", "datevalue(issuedate) = #" & countDate & "#")
        
        'DV
        '(lol)
        
        '483 Spot Checks
        '(Has table, but needs a review)
        
        .update
        .Close
    End With
fexit:
    Set RS = Nothing
    Set qdf = Nothing
    Exit Function
fErr:
    ErrHandler err, Error$, "UtilAutoCount.newCount"
    
End Function
