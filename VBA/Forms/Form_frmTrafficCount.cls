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
Option Explicit

Dim qdfWC As DAO.QueryDef
Dim qdfIC As DAO.QueryDef
Dim qdfAC As DAO.QueryDef
Dim qdfICTotals As DAO.QueryDef
Dim dateChange As Boolean

Private Sub acCount_DblClick(cancel As Integer)
    Dim oa As String: oa = "SELECT tblTraffic.ID, tblTraffic.Callsign, tbltraffic.number & '/' & tblTraffic.acType AS Type, tblTraffic.Tail, tblTraffic.countType AS [Counted As], tblTraffic.depPoint, tblTraffic.depDate, tblTraffic.Destination, tblTraffic.arrDate " & _
        "FROM tblTraffic " & _
        "WHERE ((tblTraffic.ATD is not null And tblTraffic.depPoint = DLookup('data','tblSettings','key = ""station""') And ((IIf(IsDate(tblTraffic.depDate), DateValue(tblTraffic.depDate), Null) Between CDate('" & date1 & "') And CDate('" & date2 & "')))) " & _
        "OR (tblTraffic.ATA is not null And tblTraffic.Destination = DLookup('data','tblSettings','key = ""station""') And ((IIf(IsDate(tblTraffic.arrDate), DateValue(tblTraffic.arrDate), Null) Between CDate('" & date1 & "') And CDate('" & date2 & "'))))) " & _
        "AND tblTraffic.countType Like '" & acCount.Column(0) & "' " & _
        "ORDER BY tblTraffic.arrDate DESC, tblTraffic.depDate DESC;"

    DoCmd.OpenForm "frmTrafficCountDetails", , , , , , oa

End Sub

Private Sub btnAllFlights_Click()
    DoCmd.OpenForm "checked_flts"
End Sub

Private Sub btnDeleteCount_Click()
If IsNull(shiftCount) Then Exit Sub
    Dim db As DAO.Database: Set db = CurrentDb
    db.Execute "DELETE FROM tblAnalyticsData WHERE countDate = #" & shiftCount.Column(0) & "#"
    log "Deleted " & db.RecordsAffected & " records", Me.Name & ".btnDeleteCount_Click"
    Set db = Nothing
    populate
End Sub

Private Sub btnEdit_Click()
    DoCmd.OpenForm "frmTrafficCountEdit"
End Sub

Private Sub btnNewCount_Click()
    'If Not DMax("countDay", "tblAnalytics") >= Date Then
        DoCmd.OpenForm "frmTrafficCountEntry", , , , , acDialog
        populate
    'Else
        'MsgBox "This day should already exist. If this is an error, please report it using the feedback button.", vbInformation, "AeroStat"
    'End If
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
Dim ctl As Control
    For Each ctl In Controls
        If TypeOf ctl Is ListBox Then
            ctl.Value = ""
        ElseIf TypeOf ctl Is TextBox Then
            shiftCount.SetFocus
        End If
    Next
End Sub

Private Sub Form_Load()
    Set qdfWC = CurrentDb.QueryDefs("qWeightCategory")
    Set qdfIC = CurrentDb.QueryDefs("qAnalyticsItemCountSORT")
    Set qdfAC = CurrentDb.QueryDefs("qAnalyticsACCount")
    Set qdfICTotals = CurrentDb.QueryDefs("qAnalyticsItemCountTotals")
    
    populate
End Sub

Private Sub isTransient_AfterUpdate()
    Me.lblLoadingPasses.Visible = True
    DoEvents
    qdfWC.Parameters("isTransient") = Nz(isTransient, False)
    Set lstPassCategory.Recordset = qdfWC.OpenRecordset
    Me.lblLoadingPasses.Visible = False
    Me.lblPassCat.Caption = IIf(isTransient, "Transient", "Local") & " Passes"
    DoEvents
End Sub

Private Sub qhList_DblClick(cancel As Integer)
On Error Resume Next
    DoCmd.OpenForm "quick_input", , , "ID = " & qhList.Column(0), acFormEdit
End Sub

Private Sub return_Click()
    DoCmd.Close acForm, "frmTrafficCount"
End Sub

Private Function populate()
On Error GoTo poperr
Dim RS As DAO.Recordset
Dim sc, sct, ac, qh As String
date1 = Nz(date1, Nz(DMin("countDate", "tblAnalyticsData"), DateSerial(Year(Date), 1, 1)))
date2 = Nz(date2, DMax("countDate", "tblAnalyticsData"))

On Error Resume Next
    Dim totals As Variant
    totals = sortPivotColumns("qAnalyticsItemCount", _
                        "qAnalyticsItemCountSORT", _
                        "tblAnalyticsItems", _
                        "itemName", _
                        "order", 1, _
                        date1, date2)
On Error GoTo poperr

    lstShiftCountTotals.RowSource = totals & "#" & date1 & "# And #" & date2 & "#"
    
    qdfWC.Parameters("date1") = date1
    qdfWC.Parameters("date2") = date2
    qdfWC.Parameters("isTransient") = Nz(isTransient, True)
    Set lstPassCategory.Recordset = qdfWC.OpenRecordset
    
    qdfIC.Parameters("date1") = date1
    qdfIC.Parameters("date2") = date2
    Set shiftCount.Recordset = qdfIC.OpenRecordset
    
    qdfICTotals.Parameters("date1") = date1
    qdfICTotals.Parameters("date2") = date2
    Set lstShiftCountTotals.Recordset = qdfICTotals.OpenRecordset
    
    qdfAC.Parameters("date1") = date1
    qdfAC.Parameters("date2") = date2
    Set acCount.Recordset = qdfAC.OpenRecordset
    
    If qdfIC.Fields.Count <> 0 Then
        shiftCount.ColumnCount = qdfIC.Fields.Count
        lstShiftCountTotals.ColumnCount = qdfIC.Fields.Count
    End If
    
    Dim strTotals As Variant: strTotals = Null
    Dim f: For Each f In qdfIC.Fields
        If f.Name <> "Date" Then
            strTotals = IIf(Not IsNull(strTotals), strTotals & ", sum([" & f.Name & "])", "sum([" & f.Name & "])")
        End If
    Next f
    
    strTotals = "SELECT " & strTotals & " FROM qAnalyticsItemCount WHERE countDate Between #" & date1 & "# And #" & date2 & "#"
    lstShiftCountTotals.RowSource = strTotals
    
    Dim cw As String: cw = "1"""
    Dim i: For i = 1 To qdfIC.Fields.Count + 1
        cw = cw & ";1.3"""
    Next i
    shiftCount.ColumnWidths = cw
    lstShiftCountTotals.ColumnWidths = Right(cw, Len(cw) - 3)

    'ac = "SELECT qArr.Type, qArr.Arrivals, qDep.Departures  FROM (SELECT tblTraffic.countType AS Type, Sum(tblTraffic.Number) AS Arrivals FROM tblTraffic " & _
            "WHERE (((tblTraffic.Destination)=DLookup("data","tblSettings","key = 'station'"))) " & _
            "AND ((IIf(IsDate(tblTraffic.arrDate), DateValue(tblTraffic.arrDate), Null) Between CDate('" & date1 & "') And CDate('" & date2 & "'))) " & _
            "AND ((tblTraffic.Status) Like 'Closed') GROUP BY tblTraffic.countType HAVING (((tblTraffic.countType) Is Not Null))) as qArr LEFT JOIN (SELECT tblTraffic.countType AS Type, Sum(tblTraffic.Number) AS Departures FROM tblTraffic WHERE(((tblTraffic.depPoint)=DLookup("data","tblSettings","key = 'station'"))) " & _
            "AND ((IIf(IsDate(tblTraffic.depDate), DateValue(tblTraffic.depDate), Null) Between CDate('" & date1 & "') And CDate('" & date2 & "'))) " & _
            "AND ((tblTraffic.Status) Like 'Closed' Or (tblTraffic.Status) Like 'Enroute') GROUP BY tblTraffic.countType HAVING (((tblTraffic.countType) Is Not Null))) as qDep ON qArr.Type = qDep.Type " & _
            "UNION SELECT qDep.Type, qArr.Arrivals, qDep.Departures  FROM (SELECT tblTraffic.countType AS Type, Sum(tblTraffic.Number) AS Arrivals FROM tblTraffic " & _
            "WHERE (((tblTraffic.Destination)=DLookup("data","tblSettings","key = 'station'"))) " & _
            "AND ((IIf(IsDate(tblTraffic.arrDate), DateValue(tblTraffic.arrDate), Null) Between CDate('" & date1 & "') And CDate('" & date2 & "'))) " & _
            "AND ((tblTraffic.Status) Like 'Closed')GROUP BY tblTraffic.countType HAVING (((tblTraffic.countType) Is Not Null))) as qArr RIGHT JOIN (SELECT tblTraffic.countType AS Type, Sum(tblTraffic.Number) AS Departures  FROM tblTraffic  WHERE (((tblTraffic.depPoint)=DLookup("data","tblSettings","key = 'station'")))  AND ((IIf(IsDate(tblTraffic.depDate), DateValue(tblTraffic.depDate), Null) Between CDate('" & date1 & "') And CDate('" & date2 & "')))   AND ((tblTraffic.Status) Like 'Closed' Or (tblTraffic.Status) Like 'Enroute') " & _
            "GROUP BY tblTraffic.countType  HAVING (((tblTraffic.countType) Is Not Null))) as qDep ON qDep.Type = qArr.Type   WHERE qArr.Type IS NULL;"
 
    qh = "SELECT tblTraffic.ID, tblTraffic.Callsign, tblTraffic.acType, tblTraffic.Tail, Format(IIf([Destination] Like DLookup('data','tblSettings','key = ""station""'),[arrdate],''),'dd-mmm-yy hh:nn') AS [Arrival Date (L)], Format(IIf([depPoint] Like DLookup('data','tblSettings','key = ""station""'),[depDate],''),'dd-mmm-yy hh:nn') AS [Departure Date (L)], IIf(([ATA] Between Format(ltoz('22:00'),'Short Time') And Format(ltoz('06:00'),'Short Time')) And ([Destination] Like DLookup('data','tblSettings','key = ""station""')),IIf(([ATD] Between Format(ltoz('22:00'),'Short Time') And Format(ltoz('06:00'),'Short Time')) And ([depPoint] Like DLookup('data','tblSettings','key = ""station""')),'Both','Arrival'),'Departure') AS Movement " & _
        "FROM tblTraffic " & _
        "WHERE ((Format([ATA],'Short Time') Between Format('" & LToZ("22:00") & "','Short Time') And Format('" & LToZ("06:00") & "','Short Time')) AND ((Format(tblTraffic.arrDate,'mm/dd/yyyy') Between Format('" & date1 & "','mm/dd/yyyy') And Format('" & date2 & "','mm/dd/yyyy')) And (tblTraffic.Destination)=DLookup('data','tblSettings','key = ""station""'))) " & _
        "OR ((Format([ATD],'Short Time') Between Format('" & LToZ("22:00") & "','Short Time') And Format('" & LToZ("06:00") & "','Short Time')) AND ((Format(tblTraffic.depDate,'mm/dd/yyyy')) Between Format('" & date1 & "' ,'mm/dd/yyyy') And Format('" & date2 & "','mm/dd/yyyy')) And (tblTraffic.depPoint)=DLookup('data','tblSettings','key = ""station""'));"

    'acCount.RowSource = ac
    
    qhList.RowSource = qh
    
    qhTotal = IIf(qhList.ListCount <= 0, 1, qhList.ListCount) - 1
    
    'CurrentDb.QueryDefs("qTrafficCount").sql = ac
    
    acArr = colSum(acCount, 1)
    acDep = colSum(acCount, 2)

popexit:
    Exit Function
    Resume Next
poperr:
    MsgBox Error$
   Exit Function
End Function

Private Function colSum(ByRef ctl As Control, col As Integer) As Integer
'If IsNull(ctl) Then Exit Function

    Dim i: For i = 1 To ctl.ListCount - 1
        colSum = colSum + IIf(ctl.Column(col, i) = "", 0, ctl.Column(col, i))
    Next i
    
End Function

Private Sub shiftCount_DblClick(cancel As Integer)
    If IsNull(shiftCount.Column(0)) Then Exit Sub
    DoCmd.OpenForm "frmTrafficCountEntry", , , , , acDialog, shiftCount.Column(0)
    populate
End Sub

Private Sub timeBlock_Click()
    
    Select Case timeBlock
    Case 1 'All Time
        cbYear = Null
        date1 = DMin("countDate", "tblAnalyticsData")
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

Public Function sortPivotColumns(querynameSource As String, queryname As String, SortName As String, SortColumnNameField As String, SortIndexName As String, NonPivotFieldCount As Integer, ParamArray ParamArr() As Variant) As String

' This sub goes through several steps to effectively adds an In list that sorts the 'Columns' of a crosstab query in MS Access
' 13 November 2012
' E Easterly
'
' This technique uses several components.
' 1) The original unmodified cross tab query (querynameSource)
' 2) The resulting, columns-have-been-sorted query (query)
' 3) An index table which has two columns, a numeric index used for sorting and the column name
' 4) A table or query that can be joined on the column names of the cross tab query to update the index table
'    The name of the table or query would be 'SortName'
'    The field in 'SortName' that the crosstab query columns are joined against is the 'SortColumnNameField'
'    The field in 'SortName' that has the desired order is the SortIndexName
' 5) A number which specifies the count of non-pivot/row heading columns (NonPivotFieldCount)
' 6) An optional array that contains any parameters needed for the query
'
'
'   USE:
'
'   SortPivotColumns "qryCrosstab_Initial", _
'                 "qryCrosstab_Sorted", _
'                 "tblKeyDescriptions", _
'                 "Descriptions", _
'                 "NumericIndexForSorting", _
'                  1
'
'
'
'
Dim RS As DAO.Recordset
Dim db As Database
Dim fld As DAO.Field
Dim sql As String
Dim ColumnHeading As Variant
Dim qdf As QueryDef
Dim qdfSRC As QueryDef
Dim UpdateIndexSQL As Variant
    
    DoCmd.SetWarnings False 'Turn off warnings
    
    Set db = CurrentDb
    
    Set qdfSRC = db.QueryDefs(querynameSource)
    Set qdf = db.QueryDefs(queryname)
    qdf.sql = qdfSRC.sql
    
    If Not (IsEmpty(ParamArr)) Then
        Dim i As Integer
        For i = 0 To UBound(ParamArr)
            qdf.Parameters(i) = ParamArr(i)
        Next
    End If
    
    
    ' First, get the list of fields from the query
    
    Set RS = qdf.OpenRecordset
    
    ' Then, create a temporary indexing table
    If Not IsNull(DLookup("Name", "MSysObjects", "Name='ttblSortCrosstabColumns' And Type In (1,4,6)")) Then
        db.Execute "DROP TABLE ttblSortCrosstabColumns"
    End If
    
    db.Execute "CREATE TABLE ttblSortCrosstabColumns (FieldIndex INTEGER , ColumnName TEXT(250))"
    
    ' And populate it with the current index and column names from queryname
      For Each fld In RS.Fields
        If fld.OrdinalPosition > (NonPivotFieldCount - 1) Then
            DoCmd.RunSQL "Insert into ttblSortCrosstabColumns VALUES(" & fld.OrdinalPosition & ", """ & fld.Name & """)"
        End If
      Next fld
      Set fld = Nothing
      RS.Close
      Set RS = Nothing
    
    
    ' Now, the temporary table is joined with the sort table/query and the indexes are updated
    UpdateIndexSQL = ("  UPDATE ttblSortCrosstabColumns " & _
                      "  INNER JOIN " & SortName & " ON ttblSortCrosstabColumns.ColumnName=" & SortName & "." & SortColumnNameField & _
                      "  Set ttblSortCrosstabColumns.FieldIndex = [" & SortIndexName & "]")
    DoCmd.RunSQL (UpdateIndexSQL)
    
    
    ' Then, the column headings are added to a string to prepare the In list
    sql = "SELECT ttblSortCrosstabColumns.ColumnName FROM ttblSortCrosstabColumns ORDER BY ttblSortCrosstabColumns.FieldIndex"
    Set RS = db.OpenRecordset(sql)
        RS.MoveFirst
        ColumnHeading = "'" & RS.Fields(0).Value & "'"
        sortPivotColumns = "SELECT Sum([" & RS.Fields(0).Value & "])"
        RS.MoveNext
    
        Do While Not RS.EOF
        ColumnHeading = ColumnHeading & ", '" & RS.Fields(0).Value & "'"
        sortPivotColumns = sortPivotColumns & ", Sum([" & RS.Fields(0).Value & "])"
        RS.MoveNext
        Loop
        
        sortPivotColumns = sortPivotColumns & " FROM " & queryname
    
    RS.Close
    Set RS = Nothing
    db.Execute "DROP TABLE ttblSortCrosstabColumns"
    
    Dim cs As Variant
    
    ' Set qdf = db.QueryDefs(queryname)   ' may not need this
    
    ' The query is updated with the In list
    cs = Left$(qdf.sql, Len(qdf.sql) - 3) & " In(" & ColumnHeading & ");"
    
    qdf.sql = cs
    
    
    ' Take a look at the resulting query sql by uncommenting the below section
    'Debug.Print cs
    
    DoCmd.SetWarnings True  'Turn warnings back on

End Function
'Private Sub timeToday_Click()
'    timeToday = True
'    alltime = False
'    date1 = Date
'    date2 = Date
'
'    populate
'End Sub
