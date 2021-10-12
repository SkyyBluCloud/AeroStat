VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmAllPPRs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Dim dateChange As Boolean
Dim bs As String

Private Sub btnC2imera_Click()
If ppr_list.ItemsSelected.Count = 0 Then Exit Sub
Dim PPRs() As String, idx As Integer
Dim i: For Each i In ppr_list.ItemsSelected
    ReDim Preserve PPRs(0 To idx) As String
    PPRs(idx) = ppr_list.Column(1, i)
    idx = idx + 1
Next i

C2IMERA.getCSV "'" & join(PPRs, "')) OR (((tblPPR.PPR)='") & "'"
End Sub

Private Sub btnNextDay_Click()
date1 = DateAdd("d", 1, date1)
date2 = DateAdd("d", 1, date2)
date1_Change
End Sub

Private Sub btnPrevDay_Click()
date1 = DateAdd("d", -1, date1)
date2 = DateAdd("d", -1, date2)
date1_Change
End Sub

Private Sub btnSearch_Click()
    If Not Nz(search) = "" Then
        bs = "AND ((((tblPPR.PPR) Like '*" & search & "*')) OR (((tblPPR.Callsign) Like '*" & search & "*')) OR (((tblPPR.Type) Like '*" & search & "*')) OR (((tblPPR.Tail) Like '*" & search & "*')) OR (((tblPPR.pocName) Like '*" & search & "*')) OR (((tblPPR.ctcInfo) Like '*" & search & "*')) OR (((tblPPR.Services) Like '*" & search & "*')) OR (((tblPPR.Remarks) Like '*" & search & "*'))) "
        timeBlock = 1
        timeBlock_Click
        Exit Sub
        
    Else
        bs = ""
    End If
    populate bs
End Sub

Private Sub date1_Change()
    timeBlock = ""
    If Not date2 = "" Then
        ppr_list.SetFocus
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

Private Sub date2_Change()
    timeBlock = ""
    If Not date1 = "" Then
        ppr_list.SetFocus
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

Private Sub dateCat_AfterUpdate()
populate
End Sub

Private Sub Detail_Click()
    For Each ctl In Controls
        If TypeOf ctl Is ListBox Then
            ctl.Value = ""
        End If
    Next
        
End Sub

Private Sub Form_Load()
bs = "AND (((tblTraffic.PPR) Like '*" & search & "*')) OR (((tblTraffic.Callsign) Like '*" & search & "*')) OR (((tblTraffic.Type) Like '*" & search & "*')) OR (((tblTraffic.Tail) Like '*" & search & "*')) OR (((tblTraffic.Requester) Like '*" & search & "*')) OR (((tblTraffic.ctcInfo) Like '*" & search & "*')) OR (((tblTraffic.Remarks) Like '*" & search & "*')) "
populate
End Sub

Private Sub new_Click()
    DoCmd.OpenForm "new_ppr", , , , acFormAdd, acDialog
    onPPRFormClose
End Sub

Private Sub ppr_list_DblClick(cancel As Integer)
    If Not IsNull(Me.ppr_list.Column(0)) Then
        'DoCmd.OpenForm "new_ppr", , , "ID=" & ppr_list.Column(0), acFormEdit
        DoCmd.OpenForm "new_ppr", , , "ID=" & ppr_list.Column(0), acFormEdit, acDialog
        onPPRFormClose
    End If
End Sub

Private Sub onPPRFormClose()
If Not CurrentProject.AllForms("new_ppr").IsLoaded Then Exit Sub

    If Forms!new_ppr.saveResult Then
        populate
    End If
    DoCmd.Close acForm, "new_ppr"
End Sub

Private Sub print_Click()
    populate , True
End Sub

Private Sub return_Click()
    DoCmd.Close acForm, "frmAllPPRs", acSaveNo
End Sub

Public Function populate(Optional ByVal qsearch As String, Optional ByVal prt As Boolean)
On Error GoTo errtrap
Dim RS As DAO.Recordset
Dim dc As String
dc = IIf(dateCat = 1, "arrDate", IIf(dateCat = 2, "depDate", "issueDate"))

qsearch = IIf(qsearch = "", IIf(Nz(search) = "", "", bs), qsearch)

Dim range As String
'range = "WHERE Format('" & date1 & "','mmddyy') Between Format(tblPPR.arrDate,'mmddyy') And Format(tblPPR.depDate,'mmddyy') OR Format('" & date2 & "','mmddyy') Between Format(tblPPR.arrDate,'mmddyy') And Format(tblPPR.depDate,'mmddyy') OR Format(tblPPR.arrDate,'mmddyy') Between Format('" & date1 & "','mmddyy') And Format('" & date2 & "','mmddyy')"
range = "WHERE serialDate(nz([" & dc & "],[" & IIf(dc = "arrDate", "depDate", IIf(dc = "depDate", "issueDate", "arrDate")) & "])) Between DateSerial(" & Year(date1) & "," & Month(date1) & "," & Day(date1) & ") And DateSerial(" & Year(date2) & "," & Month(date2) & "," & Day(date2) & ")"
'range = "WHERE serialDate(nz([arrDate])) Between DateSerial(" & Year(date1) & "," & Month(date1) & "," & Day(date1) & ") And DateSerial(" & Year(date2) & "," & Month(date2) & "," & Day(date2) & ")"
If dateCat = 0 Then range = "WHERE serialDate('" & date1 & "') Between serialDate(nz([arrDate])) and serialDate(nz([depDate])) OR serialDate('" & date2 & "') Between serialDate(nz([arrDate])) and serialDate(nz([depDate])) OR serialDate(nz([arrDate])) Between DateSerial(" & Year(date1) & "," & Month(date1) & "," & Day(date1) & ") And DateSerial(" & Year(date2) & "," & Month(date2) & "," & Day(date2) & ")"

Dim idx As Integer
Dim s() As String
For Each i In ppr_list.ItemsSelected
    ReDim Preserve s(0 To idx) As String
    s(idx) = ppr_list.Column(0, i)
    idx = idx + 1
Next

ppr_list.RowSource = "SELECT tblPPR.ID, tblPPR.PPR, tblPPR.Status, tblPPR.arrDate, tblPPR.depDate, tblPPR.Callsign, tblPPR.Type, tblPPR.Tail, tblPPR.issueDate, tblPPR.dvCode, tblPPR.Spot, tblPPR.Fuel, tblPPR.pocName, tblPPR.ctcInfo, tblPPR.Services, tblPPR.archive, tblPPR.depPoint, tblPPR.destination, tblPPR.remarks  " & _
        "FROM tblPPR " & range & qsearch & _
        " AND archive = " & (dateCat = 4) & IIf(dateCat = 4, " OR Status = 'Cancelled'", " AND Status <> 'Cancelled'") & " ORDER BY tblPPR.arrDate DESC , tblPPR.depDate;"
        '"ORDER BY tblPPR.arrDate DESC , tblPPR.depDate;"
        
    Set RS = CurrentDb.OpenRecordset(ppr_list.RowSource)
    numRecords.Caption = RS.RecordCount & " Total Records"
    
    If prt Then
'        If MsgBox("This will print a PPR report containing the PPRs shown. Continue?", vbQuestion + vbYesNo, "PPR") = vbYes Then
'            DoCmd.OpenReport "rPPR", , , Replace(range, "WHERE ", "") & qsearch
'        End If
        'DoCmd.OpenReport "rPPR", acViewReport, , Replace(range, "WHERE ", "") & qsearch
        
        
        DoCmd.OpenReport "rPPR", acViewReport, , , , ppr_list.RowSource
        DoEvents
        Reports!rPPR!btnClose.Visible = True
        
        Reports("rPPR").hLight = join(s, ",")
        Reports("rppr").Requery
        
'    ElseIf rs.RecordCount = 1 Then
'        DoCmd.OpenForm "new_ppr", , , "PPR = '" & rs!PPR & "'", acFormEdit
    End If
fexit:
    Exit Function
errtrap:
    MsgBox Error$
End Function

Private Sub search_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    btnSearch.SetFocus
    btnSearch_Click
End If
End Sub

Private Sub timeBlock_Click()
    
    Select Case timeBlock
    Case 1
        date1 = Format(DMin("arrdate", "tblppr"), "dd-mmm-yy")
        date2 = Format(DMax("arrdate", "tblppr"), "dd-mmm-yy")
        'date2 = Format(DateSerial(Year(Now), 12, 31), "dd-mmm-yy")
        populate
        
    Case 2
        date1 = Date
        date2 = Date
        populate
        
    Case 3
        date1 = Format("1/1/" & Year(Now), "dd-mmm-yy")
        date2 = Format("3/31/" & Year(Now), "dd-mmm-yy")
        populate
        
    Case 4
        date1 = Format("4/1/" & Year(Now), "dd-mmm-yy")
        date2 = Format("6/30/" & Year(Now), "dd-mmm-yy")
        populate
        
    Case 5
        date1 = Format("7/1/" & Year(Now), "dd-mmm-yy")
        date2 = Format("9/30/" & Year(Now), "dd-mmm-yy")
        populate
        
    Case 6
        date1 = Format("10/1/" & Year(Now), "dd-mmm-yy")
        date2 = Format("12/31/" & Year(Now), "dd-mmm-yy")
        populate
        
    End Select
End Sub
