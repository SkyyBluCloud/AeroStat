VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmReportsView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Public Sub initializeSubForms()
Dim ctl: For Each ctl In Me.Controls
    If TypeOf ctl Is SubForm Then
        ctl.SourceObject = ""
        DoEvents
        ctl.SourceObject = ctl.Tag
    End If
Next ctl
    

'With AF3616
'    .SourceObject = ""
'    DoEvents
'    .SourceObject = "Report.new3616"
'    '.Form.args = reportDate
'End With
'
'With rNOTAM
'    .SourceObject = ""
'    DoEvents
'    .SourceObject = "Report.rNOTAM"
'    .Report.RecordSource = Me.RecordSource = "SELECT tblNOTAM.* FROM tblNOTAM WHERE DateSerial(Year(tblNOTAM.startTime),Month(tblNOTAM.startTime),Day(tblNOTAM.startTime)) = #" & Parent.reportDate & "#;"
'End With
End Sub

Private Sub Form_Load()
initializeSubForms
End Sub

