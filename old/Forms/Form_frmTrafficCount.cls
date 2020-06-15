VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmTrafficCount"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Dim dateChange As Boolean

Private Sub acCount_DblClick(cancel As Integer)
    DoCmd.OpenForm "frmTrafficCountDetails", , , , , , _
        "SELECT Traffic.ID, Traffic.Callsign, Traffic.Type, Traffic.Tail, Traffic.countType AS [Counted As], Traffic.depPoint, Traffic.depDate, Traffic.Destination, Traffic.arrDate " & _
        "FROM Traffic " & _
        "WHERE ((Traffic.ATD is not null And Traffic.depPoint = dlookup('station','settings') And Traffic.depDate Between Format('" & date1 & "','mm/dd/yyyy') And  Format('" & date2 & "','mm/dd/yyyy')) " & _
        "OR (Traffic.ATA is not null And Traffic.Destination = dlookup('station','settings') And Traffic.arrDate Between Format('" & date1 & "','mm/dd/yyyy') And  Format('" & date2 & "','mm/dd/yyyy'))) " & _
        "AND Traffic.countType Like '" & acCount.Column(0) & "' " & _
        "ORDER BY Traffic.arrDate DESC, Traffic.depDate DESC;"
'    Forms!frmTrafficCountDetails!details.RowSource = _
'        "SELECT Traffic.ID, Traffic.Callsign, Traffic.Type, Traffic.Tail, Traffic.countType AS [Counted As], Traffic.depPoint, Traffic.depDate, Traffic.Destination, Traffic.arrDate " & _
'        "FROM Traffic " & _
'        "WHERE ((Traffic.ATD is not null And Traffic.depPoint = dlookup('station','settings')) OR (Traffic.ATA is not null And Traffic.Destination = dlookup('station','settings'))) " & _
'        "AND Traffic.countType Like '" & acCount.Column(0) & "' " & _
'        "AND ((Format(Nz(Traffic.arrDate),'mm/dd/yyyy') Between Format('" & date1 & "','mm/dd/yyyy') And  Format('" & date2 & "','mm/dd/yyyy')) " & _
'        "OR (Format(Nz(Traffic.depDate), 'mm/dd/yyyy') Between Format('" & date1 & "','mm/dd/yyyy') And  Format('" & date2 & "','mm/dd/yyyy'))) " & _
'        "ORDER BY Traffic.arrDate DESC, Traffic.depDate DESC;"

'FIX
'    Forms!frmTrafficCountDetails!details.RowSource = _
'        "SELECT Traffic.ID, Traffic.Callsign, Traffic.Type, Traffic.Tail, Traffic.countType AS [Counted As], Traffic.depPoint, Traffic.depDate, Traffic.Destination, Traffic.arrDate " & _
'        "FROM Traffic " & _
'        "WHERE ((Traffic.ATD is not null And Traffic.depPoint = dlookup('station','settings') And Traffic.depDate Between Format('" & date1 & "','mm/dd/yyyy') And  Format('" & date2 & "','mm/dd/yyyy')) " & _
'        "OR (Traffic.ATA is not null And Traffic.Destination = dlookup('station','settings') And Traffic.arrDate Between Format('" & date1 & "','mm/dd/yyyy') And  Format('" & date2 & "','mm/dd/yyyy'))) " & _
'        "AND Traffic.countType Like '" & acCount.Column(0) & "' " & _
'        "ORDER BY Traffic.arrDate DESC, Traffic.depDate DESC;"

End Sub

Private Sub cbYear_Change()
date1 = DateSerial(cbYear, Month(date1), Day(date1))
date2 = DateSerial(cbYear, Month(date2), Day(date2))
populate
End Sub

Private Sub date1_AfterUpdate()
    If Year(date1) <> Year(date2) Then
        cbYear = Null
    Else
        cbYear = Year(date1)
    End If
End Sub

'Private Sub alltime_Click()
'    alltime = True
'    timeToday = False
'    date1 = DMin("countDay", "tblAnalytics")
'    date2 = Date
'
'    populate
'End Sub

Private Sub date1_Change()
    timeBlock = ""
    If Not date2 = "" Then
        shiftCount.SetFocus
        populate
    End If
End Sub

Private Sub date1_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
    Case 37, 40
        date1 = DateAdd(IIf(Shift, "m", "d"), -1, date1)
        KeyCode = 0
        dateChange = True
    Case 38, 39
        date1 = DateAdd(IIf(Shift, "m", "d"), 1, date1)
        KeyCode = 0
        dateChange = True
    End Select
End Sub

Private Sub date1_LostFocus()
    If dateChange Then
        dateChange = False
        date1_Change
        date1.SetFocus
    End If
End Sub

Private Sub date2_AfterUpdate()
    If Year(date1) <> Year(date2) Then
        cbYear = Null
    Else
        cbYear = Year(date2)
    End If
End Sub

Private Sub date2_Change()
    timeBlock = ""
    If Not date1 = "" Then
        shiftCount.SetFocus
        populate
    End If
End Sub

Private Sub date2_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
    Case 37, 40
        date2 = DateAdd(IIf(Shift, "m", "d"), -1, date2)
        KeyCode = 0
        dateChange = True
    Case 38, 39
        date2 = DateAdd(IIf(Shift, "m", "d"), 1, date2)
        KeyCode = 0
        dateChange = True
    End Select
End Sub

Private Sub date2_LostFocus()
    If dateChange Then
        dateChange = False
        date2_Change
        date2.SetFocus
    End If
End Sub

Private Sub Detail_Click()
    For Each ctl In Controls
        If TypeOf ctl Is ListBox Then
            ctl.Value = ""
        ElseIf TypeOf ctl Is TextBox Then
            shiftCount.SetFocus
        End If
    Next
End Sub

Private Sub Form_Load()
    populate
End Sub

Private Sub qhList_DblClick(cancel As Integer)
On Error Resume Next
    DoCmd.OpenForm "quick_input", , , "ID = " & qhList.Column(0), acFormEdit
End Sub

Private Sub return_Click()
    DoCmd.Close acForm, "frmTrafficCount", acSaveYes
End Sub

Private Sub Command4_Click()
    'If Not DMax("countDay", "tblAnalytics") >= Date Then
        DoCmd.OpenForm "frmTrafficCountEntry", , , , acFormAdd
    'Else
        'MsgBox "This day should already exist. If this is an error, please report it using the feedback button.", vbInformation, "AeroStat"
    'End If
End Sub

Private Function populate()
On Error GoTo poperr
Dim rs As DAO.Recordset
Dim sc As String
Dim sct As String
Dim ac As String
Dim qh As String

date1 = Nz(date1, DMin("countDay", "tblAnalytics"))
date2 = Nz(date2, DMax("countDay", "tblAnalytics"))

    sc = "SELECT  tblAnalytics.ID, tblAnalytics.countDay, tblAnalytics.IFE, tblAnalytics.GE, tblAnalytics.DD1801, tblAnalytics.DD1801Stereo, tblAnalytics.transientArr, tblAnalytics.DVs, tblAnalytics.hazCargo, tblAnalytics.Inspections, tblAnalytics.Checks, tblAnalytics.AF483, tblAnalytics.BASH, tblAnalytics.fireStreamers, tblAnalytics.fireBanger, tblAnalytics.fireLive, tblAnalytics.Depredation " & _
        "FROM tblAnalytics " & _
        "WHERE tblAnalytics.countDay Between Format('" & date1 & "','mm/dd/yyyy') And Format('" & date2 & "','mm/dd/yyyy') " & _
        "ORDER BY tblAnalytics.countDay DESC;"
        
    sct = "SELECT Sum(tblAnalytics.IFE) AS SumOfIFE, Sum(tblAnalytics.GE) AS SumOfGE, Sum(tblAnalytics.DD1801) AS SumOfDD1801, Sum(tblAnalytics.DD1801Stereo) AS SumOfDD1801Stereo, Sum(tblAnalytics.transientArr) AS SumOftransientArr, Sum(tblAnalytics.DVs) AS SumOfDVs, Sum(tblAnalytics.hazCargo) AS SumOfhazCargo, Sum(tblAnalytics.Inspections) AS SumOfInspections, Sum(tblAnalytics.Checks) AS SumOfChecks, Sum(tblAnalytics.AF483) AS SumOfAF483, Sum(tblAnalytics.BASH) AS SumOfBASH, Sum(tblAnalytics.fireStreamers) AS SumOffireStreamers, Sum(tblAnalytics.fireBanger) AS SumOffireBanger, Sum(tblAnalytics.fireLive) AS SumOffireLive, Sum(tblAnalytics.Depredation) AS SumOfDepredation " & _
        "FROM tblAnalytics " & _
        "WHERE tblAnalytics.countDay Between Format('" & date1 & "','mm/dd/yyyy') And Format('" & date2 & "','mm/dd/yyyy');"

'This counts INDIVIDUAL records
'    ac = "SELECT qArr.Type, qArr.Arrivals, qDep.Departures " & _
'            "FROM (SELECT Traffic.countType AS Type, Count(Traffic.ATA) AS Arrivals FROM Traffic WHERE (((Traffic.Destination)=DLookUp('station','settings')) AND ((Format(Traffic.arrDate,'mm/dd/yyyy') Between Format('" & date1  & "','mm/dd/yyyy') And  Format('" & date2  & "','mm/dd/yyyy')))) AND ((Traffic.Status) Like 'Closed') GROUP BY Traffic.countType HAVING (((Traffic.countType) Is Not Null))) as qArr " & _
'            "LEFT JOIN (SELECT Traffic.countType AS Type, Count(Traffic.ATD) AS Departures FROM Traffic WHERE (((Traffic.depPoint)=DLookUp('station','settings')) AND ((Format(traffic.depDate,'mm/dd/yyyy')) Between Format('" & date1  & "','mm/dd/yyyy') And Format('" & date2  & "','mm/dd/yyyy')) AND ((Traffic.Status) Like 'Closed' Or (Traffic.Status) Like 'Enroute')) GROUP BY Traffic.countType HAVING (((Traffic.countType) Is Not Null))) as qDep " & _
'            "ON qArr.Type = qDep.Type " & _
'        "UNION " & _
'            "SELECT qDep.Type, qArr.Arrivals, qDep.Departures " & _
'            "FROM (SELECT Traffic.countType AS Type, Count(Traffic.ATA) AS Arrivals FROM Traffic WHERE (((Traffic.Destination)=DLookUp('station','settings')) AND ((Format(Traffic.arrDate,'mm/dd/yyyy') Between Format('" & date1  & "','mm/dd/yyyy') And  Format('" & date2  & "','mm/dd/yyyy')))) AND ((Traffic.Status) Like 'Closed') GROUP BY Traffic.countType HAVING (((Traffic.countType) Is Not Null))) as qArr " & _
'            "RIGHT JOIN (SELECT Traffic.countType AS Type, Count(Traffic.ATD) AS Departures FROM Traffic WHERE (((Traffic.depPoint)=DLookUp('station','settings')) AND ((Format(traffic.depDate,'mm/dd/yyyy')) Between Format('" & date1  & "','mm/dd/yyyy') And Format('" & date2  & "','mm/dd/yyyy')) AND ((Traffic.Status) Like 'Closed' Or (Traffic.Status) Like 'Enroute')) GROUP BY Traffic.countType HAVING (((Traffic.countType) Is Not Null))) as qDep " & _
'            "ON qDep.Type = qArr.Type " & _
'        "WHERE qArr.Type IS NULL;"

'This counts NUMBER of acft per flight plan | 26 Jul: Removed Format() from arrDate and depDate
    ac = "SELECT qArr.Type, qArr.Arrivals, qDep.Departures " & _
            "FROM (SELECT Traffic.countType AS Type, Sum(Traffic.Number) AS Arrivals FROM Traffic WHERE (((Traffic.Destination)=DLookUp('station','settings')) AND ((Traffic.arrDate Between Format('" & date1 & "','mm/dd/yyyy') And  Format('" & date2 & "','mm/dd/yyyy')))) AND ((Traffic.Status) Like 'Closed') GROUP BY Traffic.countType HAVING (((Traffic.countType) Is Not Null))) as qArr " & _
            "LEFT JOIN (SELECT Traffic.countType AS Type, Sum(Traffic.Number) AS Departures FROM Traffic WHERE (((Traffic.depPoint)=DLookUp('station','settings')) AND ((Traffic.depDate) Between Format('" & date1 & "','mm/dd/yyyy') And Format('" & date2 & "','mm/dd/yyyy')) AND ((Traffic.Status) Like 'Closed' Or (Traffic.Status) Like 'Enroute')) GROUP BY Traffic.countType HAVING (((Traffic.countType) Is Not Null))) as qDep " & _
            "ON qArr.Type = qDep.Type " & _
        "UNION " & _
            "SELECT qDep.Type, qArr.Arrivals, qDep.Departures " & _
            "FROM (SELECT Traffic.countType AS Type, Sum(Traffic.Number) AS Arrivals FROM Traffic WHERE (((Traffic.Destination)=DLookUp('station','settings')) AND ((Traffic.arrDate Between Format('" & date1 & "','mm/dd/yyyy') And  Format('" & date2 & "','mm/dd/yyyy')))) AND ((Traffic.Status) Like 'Closed') GROUP BY Traffic.countType HAVING (((Traffic.countType) Is Not Null))) as qArr " & _
            "RIGHT JOIN (SELECT Traffic.countType AS Type, Sum(Traffic.Number) AS Departures FROM Traffic WHERE (((Traffic.depPoint)=DLookUp('station','settings')) AND ((Traffic.depDate) Between Format('" & date1 & "','mm/dd/yyyy') And Format('" & date2 & "','mm/dd/yyyy')) AND ((Traffic.Status) Like 'Closed' Or (Traffic.Status) Like 'Enroute')) GROUP BY Traffic.countType HAVING (((Traffic.countType) Is Not Null))) as qDep " & _
            "ON qDep.Type = qArr.Type " & _
        "WHERE qArr.Type IS NULL;"
        
'    qh = "SELECT Traffic.ID, Traffic.Callsign, Traffic.Type, Traffic.Tail, Format(IIf([Destination] Like DLookUp('station','settings'),[arrdate],''),'dd-mmm-yy hh:nn') AS [Arrival Date (L)], Format(IIf([depPoint] Like DLookUp('station','settings'),[depDate],''),'dd-mmm-yy hh:nn') AS [Departure Date (L)], IIf(([ATA] Between Format(ltoz('22:00'),'Short Time') And Format(ltoz('06:00'),'Short Time')) And ([Destination] Like DLookUp('station','settings')),IIf(([ATD] Between Format(ltoz('22:00'),'Short Time') And Format(ltoz('06:00'),'Short Time')) And ([depPoint] Like DLookUp('station','settings')),'Both','Arrival'),'Departure') AS Movement " & _
'        "FROM Traffic " & _
'        "WHERE (((Traffic.ATA) Between Format(ltoz('22:00'),'Short Time') And Format(ltoz('06:00'),'Short Time')) AND ((Format(Traffic.arrDate,'mm/dd/yyyy') Between Format('" & date1 & "','mm/dd/yyyy') And Format('" & date2 & "','mm/dd/yyyy')) And (Traffic.Destination)=DLookUp('station','settings'))) OR (((Traffic.ATD) Between Format(ltoz('22:00'),'Short Time') And Format(ltoz('06:00'),'Short Time')) AND ((Format(Traffic.depDate,'mm/dd/yyyy')) Between Format('" & date1 & "' ,'mm/dd/yyyy') And Format('" & date2 & "','mm/dd/yyyy')) And (Traffic.depPoint)=DLookUp('station','settings'));"
        
    qh = "SELECT Traffic.ID, Traffic.Callsign, Traffic.Type, Traffic.Tail, Format(IIf([Destination] Like DLookUp('station','settings'),[arrdate],''),'dd-mmm-yy hh:nn') AS [Arrival Date (L)], Format(IIf([depPoint] Like DLookUp('station','settings'),[depDate],''),'dd-mmm-yy hh:nn') AS [Departure Date (L)], IIf(([ATA] Between Format(ltoz('22:00'),'Short Time') And Format(ltoz('06:00'),'Short Time')) And ([Destination] Like DLookUp('station','settings')),IIf(([ATD] Between Format(ltoz('22:00'),'Short Time') And Format(ltoz('06:00'),'Short Time')) And ([depPoint] Like DLookUp('station','settings')),'Both','Arrival'),'Departure') AS Movement " & _
        "FROM Traffic " & _
        "WHERE ((Format([ATA],'Short Time') Between Format('" & LToZ("22:00") & "','Short Time') And Format('" & LToZ("06:00") & "','Short Time')) AND ((Format(Traffic.arrDate,'mm/dd/yyyy') Between Format('" & date1 & "','mm/dd/yyyy') And Format('" & date2 & "','mm/dd/yyyy')) And (Traffic.Destination)=DLookUp('station','settings'))) " & _
        "OR ((Format([ATD],'Short Time') Between Format('" & LToZ("22:00") & "','Short Time') And Format('" & LToZ("06:00") & "','Short Time')) AND ((Format(Traffic.depDate,'mm/dd/yyyy')) Between Format('" & date1 & "' ,'mm/dd/yyyy') And Format('" & date2 & "','mm/dd/yyyy')) And (Traffic.depPoint)=DLookUp('station','settings'));"
    
        
    shiftCount.RowSource = sc

    shiftCountTotal.RowSource = sct

    acCount.RowSource = ac
    
    qhList.RowSource = qh
    
    qhTotal = IIf(qhList.ListCount <= 0, 1, qhList.ListCount) - 1
    
    CurrentDb.QueryDefs("qTrafficCount").sql = ac
    acArr = DSum("qArr.Arrivals", "qTrafficCount")
    acDep = DSum("qDep.Departures", "qTrafficCount")

popexit:
    Exit Function
    Resume Next
poperr:
    MsgBox Error$
   
End Function

Private Sub shiftCount_DblClick(cancel As Integer)
    If IsNull(shiftCount.Column(0)) Then Exit Sub
    DoCmd.OpenForm "frmTrafficCountEntry", , , "ID=" & shiftCount.Column(0), acFormEdit
End Sub

Private Sub timeBlock_Click()
    
    Select Case timeBlock
    Case 1 'All Time
        cbYear = Null
        date1 = DMin("countDay", "tblAnalytics")
        date2 = Date
        populate
        
    Case 2 'Today
        cbYear = Year(Date)
        date1 = Format(Date, "dd-mmm-yy")
        date2 = date1
        populate
        
    Case 3 'Q1
        date1 = Format("1/1/" & Nz(Nz(cbYear, Year(Date)), Year(Date)), "dd-mmm-yy")
        date2 = Format("3/31/" & Nz(cbYear, Year(Date)), "dd-mmm-yy")
        populate
        
    Case 4 'Q2
        date1 = Format("4/1/" & Nz(cbYear, Year(Date)), "dd-mmm-yy")
        date2 = Format("6/30/" & Nz(cbYear, Year(Date)), "dd-mmm-yy")
        populate
        
    Case 5 'Q3
        date1 = Format("7/1/" & Nz(cbYear, Year(Date)), "dd-mmm-yy")
        date2 = Format("9/30/" & Nz(cbYear, Year(Date)), "dd-mmm-yy")
        populate
        
    Case 6 'Q4
        date1 = Format("10/1/" & Nz(cbYear, Year(Date)), "dd-mmm-yy")
        date2 = Format("12/31/" & Nz(cbYear, Year(Date)), "dd-mmm-yy")
        populate
        
    Case 7 'Year
        date1 = Format("1/1/" & Nz(cbYear, Year(Date)), "dd-mmm-yy")
        date2 = Format("12/31/" & Nz(cbYear, Year(Date)), "dd-mmm-yy")
        populate
        
    End Select
End Sub

'Private Sub timeToday_Click()
'    timeToday = True
'    alltime = False
'    date1 = Date
'    date2 = Date
'
'    populate
'End Sub
