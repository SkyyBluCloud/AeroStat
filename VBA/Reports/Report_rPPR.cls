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
Dim doOrderArrDate, doOrderDepDate, doOrderPPR, obArrDate, obDepDate, obPPR As Boolean


Private Sub arrdate_Label_Click()
doOrderArrDate = True
obArrDate = Not obArrDate
refreshOrder 1
End Sub

Private Sub depdate_Label_Click()
doOrderDepDate = True
obDepDate = Not obDepDate
refreshOrder 2
End Sub

Private Sub email_Click()
On Error GoTo errtrap
'Dim oApp As New Outlook.Application
'Dim oEmail As Outlook.MailItem
Dim file, subj, body As String

    file = DLookup("dbroot", "tblSettings") & "\PPR Reports\" & Replace(Caption, ":", "") & ".pdf"
    'DoCmd.OutputTo acOutputReport, Me.Name, acFormatPDF, file, False
    
    subj = DLookup("pprEmailSubject", "tblsettings")
    subj = Replace(subj, "[ppr]", PPR)
    subj = Replace(subj, "[callsign]", UCase(Callsign))
    subj = Replace(subj, "[type]", UCase(Me.Type))
    subj = Replace(subj, "[arrdate]", UCase(Format(arrDate, "d mmm yy")))
    
    body = DLookup("pprEmailBody", "tblsettings")
    
    DoCmd.SendObject acSendReport, Me.Name, acFormatPDF, , , , subj, body, True
    DoCmd.Close acReport, Me.Name
'
'    Set oEmail = oApp.CreateItem(olMailItem)
'    With oEmail
'        .Subject = Caption
'        .Body = "Attached is the PPR report, current as of " & Format(Now(), "dd-mmm-yy hh:nn L")
'        .Attachments.add file
'    End With

sExit:
    Exit Sub
errtrap:
    ErrHandler err, Error$, Me.Name
End Sub

Private Sub PPR_Label_Click()
doOrderPPR = True
obPPR = Not obPPR
refreshOrder 3
End Sub

Private Sub print_Click()
On Error Resume Next
    DoCmd.RunCommand acCmdPrint
End Sub

Private Sub refreshOrder(Optional priority As Integer)
'obArrDate, obDepDate, obPPR
Dim order As String

If Nz(priority) <> "" Then
    Select Case priority
    Case 1
        order = "arrdate" & IIf(obArrDate, " DESC", "")
    Case 2
        order = "depDate" & IIf(obDepDate, " DESC", "")
    Case 3
        order = "PPR" & IIf(obPPR, " DESC", "")
    End Select
End If

If doOrderArrDate And priority <> 1 Then
    order = order & ", arrdate" & IIf(obArrDate, " DESC", "")
End If

If doOrderDepDate And priority <> 2 Then
    order = order & ", depDate" & IIf(obDepDate, " DESC", "")
End If

If doOrderPPR And priority <> 3 Then
    order = order & ", PPR" & IIf(obPPR, " DESC", "")
End If

Me.OrderBy = order
Me.OrderByOn = True
    
End Sub

Private Sub Report_Load()
    Caption = "PPR Report " & Format(Now, "dd-mmm-yy hhnn(L)")
    If Nz(OpenArgs) <> "" Then Me.RecordSource = OpenArgs
End Sub


