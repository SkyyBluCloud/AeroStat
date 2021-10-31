VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frm483view"
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
'This is the view to be used in the view control of frm483
Dim orig As String

Private Sub btnSearch_Click()
Dim q As String

q = "WHERE (((tbl483.rankID) Like '*" & search & "*') OR ((tbl483.AF483) Like '*" & search & "*') OR ((tbl483.lastName) Like '*" & search & "*') OR ((tbl483.firstName) Like '*" & search & "*') OR ((tbl483.midInit) Like '*" & search & "*') OR ((tbl483.Unit) Like '*" & search & "*') OR ((tbl483.issueDate) Like '*" & search & "*') OR ((tbl483.issuedBy) Like '*" & search & "*') OR ((tbl483.CMA) Like '*" & search & "*'))"
If Nz(search) = "" Then
    list483.RowSource = orig
Else
    list483.RowSource = Replace(orig, "WHERE ((Left([af483],2))=[Forms]![frm483]![view]![series]);", "")
    list483.RowSource = list483.RowSource & q
End If
    
End Sub

Private Sub Detail_Click()
Parent.Detail_Click
End Sub

Private Sub Form_Load()
On Error Resume Next
    If Parent.srs = "" Then Parent.srs = "18"
    series = Parent.srs
    list483.Requery
    list483.SetFocus
    list483.ListIndex = Parent.sel
    orig = list483.RowSource
End Sub

Private Sub list483_DblClick(cancel As Integer)
If list483.ItemsSelected.Count <> 1 Then Exit Sub
    Parent.sel = list483.ListIndex
    Parent!back.Visible = True
    Parent!back.Tag = "frm483view"
    Parent.changeView "frm483entry", "AF483 = '" & list483.Column(0) & "'", list483.Column(0)
    
'    Filter = "AF483 = '" & list483.Column(0) & "'"
'    DoCmd.OpenForm "frm483entry", , , "AF483 = '" & list483.Column(0) & "'", acFormEdit
'    view!Checks.RowSource = "SELECT ID, checkDate, compliant, reason, details, by FROM tbl483CompCheck WHERE AF483 = '" & list483.Column(0) & "' ORDER BY checkDate DESC"

End Sub

Private Sub new_Click()
    If DLookup("authlevel", "tbluserauth", "username = '" & Util.getUser & "'") > 4 Then
        MsgBox "Access denied.", vbCritical, "ADPM"
        Exit Sub
    End If
    Parent.changeView "frm483entry"
    Parent!back.Visible = True
    Parent!back.Tag = "frm483view"
End Sub

Private Sub search_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        Dirty = False
        btnSearch_Click
    End If
End Sub

Private Sub series_Change()
list483.Requery
Parent.srs = series
End Sub
