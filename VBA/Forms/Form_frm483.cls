VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frm483"
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
Public srs As String
Public sel As Integer

Public Sub back_Click()
view.SourceObject = back.Tag
view.SetFocus
back.Visible = False
End Sub

Public Sub Detail_Click()
For Each ctl In view.Controls
If TypeOf ctl Is ListBox Then ctl.Value = ""
Next
End Sub

Public Function changeView(ByVal frm As String, Optional ByVal args As String, Optional args2 As Variant = 0)
    view.SourceObject = frm
    view.Form.Filter = args
    view.Form.FilterOn = Not IsNull(args)
    
    If frm = "frm483entry" Then
        view!Checks.RowSource = "SELECT ID, checkDate, compliant, reason, details, by FROM tbl483CompCheck WHERE AF483 = '" & Nz(args2) & "' ORDER BY checkDate DESC"
    End If
    
End Function

Private Sub pov_Click()
changeView "frmPOVPass"
back.Tag = "frm483view"
back.Visible = True
End Sub
