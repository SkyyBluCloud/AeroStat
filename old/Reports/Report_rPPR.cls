VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Report_rPPR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Dim obArrDate, obDepDate, obPPR As Boolean


Private Sub arrdate_Label_Click()
Me.OrderBy = "arrdate" & IIf(obArrDate, " DESC", "")
obArrDate = Not obArrDate
Me.OrderByOn = True
End Sub

Private Sub depdate_Label_Click()
Me.OrderBy = "depDate" & IIf(obDepDate, " DESC", "")
obDepDate = Not obDepDate
Me.OrderByOn = True
End Sub

Private Sub email_Click()
On Error GoTo errTrap
'Dim oApp As New Outlook.Application
'Dim oEmail As Outlook.MailItem
Dim file As String

    file = DLookup("driveprefix", "settings") & "7. PPR and Quiet Hr Log Tracker\PPR Reports\" & Replace(Caption, ":", "") & ".pdf"
    'DoCmd.OutputTo acOutputReport, Me.Name, acFormatPDF, file, False
    DoCmd.SendObject acSendReport, Me.Name, acFormatPDF, , , , Caption, "Attached is the requested PPR Report, current as of " & Format(Now, "dd-mmm-yy hhnn(L)") & "." & vbCrLf & vbCrLf & "v/r" & vbCrLf, True
    DoCmd.Close acReport, Me.Name, acSaveNo
'
'    Set oEmail = oApp.CreateItem(olMailItem)
'    With oEmail
'        .Subject = Caption
'        .Body = "Attached is the PPR report, current as of " & Format(Now(), "dd-mmm-yy hh:nn L")
'        .Attachments.add file
'    End With

sexit:
    Exit Sub
errTrap:
    errHandler err, Error$, Me.Name
End Sub

Private Sub PPR_Label_Click()
Me.OrderBy = "PPR" & IIf(obPPR, " DESC", "")
obPPR = Not obPPR
Me.OrderByOn = True
End Sub

Private Sub print_Click()
On Error Resume Next
    DoCmd.RunCommand acCmdPrint
End Sub

Private Sub Report_Load()
    Caption = "PPR Report " & Nz(OpenArgs) & Format(Now, "dd-mmm-yy hhnn(L)")
End Sub


