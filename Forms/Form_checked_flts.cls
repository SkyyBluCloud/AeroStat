VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_checked_flts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'   AeroStat: Airfield Management(TM)
'
'
'    Copyright 2017-2020 Shane J. DeSouza(SkyyBluCloud)
'
'   Licensed under the Apache License, Version 2.0 (the "License");
'   you may not use this file except in compliance with the License.
'   You may obtain a copy of the License at
'
'       http://www.apache.org/licenses/LICENSE-2.0
'
'   Unless required by applicable law or agreed to in writing, software
'   distributed under the License is distributed on an "AS IS" BASIS,
'   WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
'   See the License for the specific language governing permissions and
'   limitations under the License.

Option Compare Database
Dim bs As String

Private Sub btnSearch_Click()
On Error GoTo errTrap
    If Not search = "" Then
        If InStr(1, search, "=") > 0 Then
            Dim sField As String
            Dim idx As Integer
            Dim fld As String
            Dim prm As String
            idx = InStr(1, search, "=")
            fld = Left(search, idx - 1)
            param = Mid(search, idx + 1)
            bs = "AND (((traffic." & fld & ") Like '*" & param & "*'))"
        Else
            bs = "AND ((((traffic.PPR) Like '*" & search & "*')) OR (((traffic.Callsign) Like '*" & search & "*')) OR (((traffic.Type) Like '*" & search & "*')) OR (((traffic.Tail) Like '*" & search & "*')) OR (((traffic.Requester) Like '*" & search & "*')) OR (((traffic.ctcInfo) Like '*" & search & "*')) OR (((traffic.Remarks) Like '*" & search & "*')))"
        End If
    Else
        bs = ""
    End If
    populate , , bs
sexit:
    Exit Sub
errTrap:
    MsgBox Error$
    Resume Next
End Sub

Private Sub btnSearchClear_Click()
    search = ""
    populate
End Sub

Private Sub customDate1_Change()
    If Not customDate2 = "" Then
        optDate = 3
        done.SetFocus
        populate
        optDate.SetFocus
    End If
End Sub

Private Sub customDate2_Change()
    If Not customDate1 = "" Then
        optDate = 3
        done.SetFocus
        populate
        optDate.SetFocus
    End If
End Sub

Private Sub Detail_Click()
    done = ""
End Sub

Private Sub done_DblClick(cancel As Integer)

    If Not IsNull(Me.done.Column(0)) Then
        DoCmd.OpenForm "quick_input", , , "ID=" & Me.done.Column(0), acFormEdit
        For Each ctl In Forms!quick_input.Controls
            If ctl.Name Like "*ctime*" Then ctl.Visible = True
        Next
    End If
    
check_out_Click_err:
    Exit Sub
End Sub

Private Sub Form_Load()
Dim rs As DAO.Recordset
Direction = "Is Not Null"
bs = "AND (((traffic.PPR) Like '*" & search & "*')) OR (((traffic.Callsign) Like '*" & search & "*')) OR (((traffic.Type) Like '*" & search & "*')) OR (((traffic.Tail) Like '*" & search & "*')) OR (((traffic.Requester) Like '*" & search & "*')) OR (((traffic.ctcInfo) Like '*" & search & "*')) OR (((traffic.Remarks) Like '*" & search & "*')) "
populate
    Set rs = CurrentDb.OpenRecordset("SELECT Traffic.Callsign, Traffic.Status, Traffic.depPoint, Traffic.Destination, Traffic.ATD FROM Traffic WHERE Traffic.Status Like 'Enroute' AND Traffic.depPoint Not Like Traffic.Destination AND Traffic.Destination not Like DLookUp('Station','settings','ID=1')")
    If Not rs.EOF Then rs.MoveLast
    Do Until rs.BOF
    If DateAdd("n", 30, rs!ATD) < LToZ(Now()) Then
            rs.edit
            rs!Status = "Closed"
            rs.update
        End If
        rs.MovePrevious
    Loop
    rs.Close
    Set rs = Nothing
    
    
End Sub

Private Sub optDate_Click()
    populate
End Sub

Private Sub return_Click()
    DoCmd.Close acForm, "checked_flts", acSaveYes
End Sub

Private Function populate(Optional Direction As String, Optional qstatus As String, Optional qsearch As String)
Dim qaf As String
Dim rs As DAO.Recordset

    Direction = IIf(Direction = "", _
        IIf(sortDir = "All", "Is Not Null", "Like '" & sortDir & "'"), Direction)
    
    qstatus = IIf(qstatus = "", IIf(Not sts = "All", "AND ((Traffic.Status) Like '" & sts & "' )", ""), qstatus)
    
    qsearch = IIf(qsearch = "", IIf(Nz(search) = "", "", bs), qsearch)

    qaf = "SELECT Traffic.ID, IIf([depPoint]=DLookUp('Station','settings'),IIf([Destination]=[depPoint],'Local','Departure'),'Arrival') AS Direction, Traffic.DOF, Traffic.flightRule, Traffic.Callsign, Traffic.Type, Traffic.Tail, Traffic.depPoint, Traffic.flightLvl, Traffic.ETD, Traffic.ATD, Traffic.ETE, Traffic.Destination, Traffic.cETA, Traffic.ATA, Traffic.Status, Traffic.TWR_In, Traffic.TWR_Out " & _
    "FROM traffic " & _
    "WHERE (((IIf([depPoint]=DLookUp('Station','settings'),IIf([Destination]=[depPoint],'Local','Departure'),'Arrival')) " & Direction & ") " & _
    qstatus & _
    "AND ((Traffic.depPoint) Not Like DLookUp('Station','settings','ID=1')) " & _
    "AND ((Traffic.Destination) Like DLookUp('Station','settings','ID=1')) " & _
    "AND ((Format((DateAdd('h',DLookUp('Timezone','settings'),[dof]+IIf([atd] Is Null,[etd],[atd]))+[ete]),'mmddyy')) [magic])) " & qsearch & _
    "OR (((IIf([depPoint]=DLookUp('Station','settings'),IIf([Destination]=[depPoint],'Local','Departure'),'Arrival')) " & Direction & ") " & _
    qstatus & _
    "AND ((Traffic.depPoint) Like DLookUp('Station','settings','ID=1')) " & _
    "AND ((Traffic.Destination) Not Like DLookUp('Station','settings','ID=1')) " & _
    "AND ((Format(DateAdd('h',DLookUp('Timezone','settings'),[dof]+[etd]),'mmddyy')) [magic])) " & qsearch & _
    "OR (((IIf([depPoint]=DLookUp('Station','settings'),IIf([Destination]=[depPoint],'Local','Departure'),'Arrival')) " & Direction & ") " & _
    qstatus & _
    "AND ((Traffic.depPoint) Like DLookUp('Station','settings','ID=1')) " & _
    "AND ((Traffic.Destination) Like DLookUp('Station','settings','ID=1')) " & _
    "AND ((Format((DateAdd('h',DLookUp('Timezone','settings'),[dof]+IIf([atd] Is Null,[etd],[atd]))+[ete]),'mmddyy')) [magic])) " & qsearch & _
    "ORDER BY Traffic.DOF DESC, IIf(Traffic.ATD Is Null,Traffic.ETD,Traffic.ATD) DESC;"
    '"ORDER BY IIf(IIf([depPoint]=DLookUp('Station','settings'),IIf([Destination]=[depPoint],'Local','Departure'),'Arrival') Not Like 'Local',IIf(IIf([depPoint]=DLookUp('Station','settings'),IIf([Destination]=[depPoint],'Local','Departure'),'Arrival') Like 'Arrival',Traffic.arrDate,Traffic.depDate),Traffic.depDate) DESC;"

    Select Case optDate
        Case 1
            qaf = Replace(qaf, "[magic]", "Between Format([thedate],'mmddyy') And Format([thedate2],'mmddyy')")
            qaf = Replace(qaf, "[thedate]", "'" & Date & "'")
            qaf = Replace(qaf, "[thedate2]", "'" & Date & "'")
            done.RowSource = qaf
            
        Case 2
            qaf = Replace(qaf, "[magic]", "Is Not Null")
            done.RowSource = qaf
            
        Case 3
            qaf = Replace(qaf, "[magic]", "Between Format([thedate],'mmddyy') And Format([thedate2],'mmddyy')")
            qaf = Replace(qaf, "[thedate]", "'" & customDate1 & "'")
            qaf = Replace(qaf, "[thedate2]", "'" & customDate2 & "'")
            done.RowSource = qaf
            
    End Select
    optDate.Requery
    Set rs = CurrentDb.OpenRecordset(qaf)
    numFlights.Caption = rs.RecordCount & " Total Flights"
    Set rs = Nothing
    
End Function

Private Sub sortDir_AfterUpdate()
Dim sd As String
    If sortDir = "All" Then
        sd = "Is Not Null"
    Else
        sd = "Like '" & sortDir & "'"
    End If
    populate sd
End Sub

Private Sub sts_AfterUpdate()
Dim qs As String
    If Not sts = "All" Then qs = "AND ((Traffic.Status) Like '" & sts & "') " Else qs = ""
    populate , qs
End Sub
