VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frm483entry"
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

Private Sub CMA_Click()
Form_Current
log CMA, Me.Name & "CMA_Click"
End Sub

Private Sub done_Click()
    Parent.back_Click
End Sub

Private Sub Form_Current()
    Select Case CMA
        Case True
            CMA.Caption = "CMA"
        Case False
            CMA.Caption = "Ramp Only"
    End Select
End Sub

