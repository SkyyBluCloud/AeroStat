VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmADPM"
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
Dim isLoggedIn As Boolean

Private Sub tabCtl_Change()
On Error GoTo errtrap
    If tabCtl = 1 And Not isLoggedIn Then
        If MsgBox("A SharePoint connection will be attempted. Continue?", vbQuestion + vbYesNo, "ADTP") = vbYes Then
            DoCmd.OpenForm "frmLoading"
            Forms!frmLoading!loadingText.Caption = "Logging into SharePoint..."
            DoEvents
            frmPOVPass.SourceObject = "frmPOVPass"
            DoCmd.Close acForm, "frmLoading"
            isLoggedIn = True
        Else
            tabCtl = 0
        End If
    End If
sExit:
    Exit Sub
errtrap:
    
End Sub
