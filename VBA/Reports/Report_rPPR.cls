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
Public hLight As String 'CSV

Private Sub arrdate_Label_Click()
doOrderArrDate = True
obArrDate = Not obArrDate
refreshOrder 1
End Sub

Private Sub btnClose_Click()
DoCmd.Close acReport, Me.Name
End Sub

Private Sub btnEmail_Click()
On Error GoTo errtrap
'Dim oApp As New Outlook.Application
'Dim oEmail As Outlook.MailItem

'    file = DLookup("dbroot", "tblSettings") & "\PPR Reports\" & Replace(Caption, ":", "") & ".pdf"
'    'DoCmd.OutputTo acOutputReport, Me.Name, acFormatPDF, file, False
'
'    subj = DLookup("pprEmailSubject", "tblsettings")
'    subj = Replace(subj, "[ppr]", PPR)
'    subj = Replace(subj, "[callsign]", UCase(Callsign))
'    subj = Replace(subj, "[type]", UCase(Me.Type))
'    subj = Replace(subj, "[arrdate]", UCase(Format(arrDate, "d mmm yy")))
'
'    body = DLookup("pprEmailBody", "tblsettings")
'
'    DoCmd.SendObject acSendReport, Me.Name, acFormatPDF, , , , subj, body, True
'    DoCmd.Close acReport, Me.Name
    
    
    
    If MsgBox("Email request?", vbQuestion + vbYesNo, "PPR") = vbYes Then
        Dim subj, body As String
        Dim RS As DAO.Recordset: Set RS = CurrentDb.OpenRecordset(Me.RecordSource)
        With RS
            Do While Not .EOF
                If doHighlight(!ID) Then Exit Do
                .MoveNext
            Loop
            If .EOF Then .MoveFirst
        
            subj = DLookup("data", "tblSettings", "key = ""pprEmailSubject""")
            subj = Replace(subj, "[ppr]", !PPR)
            subj = Replace(subj, "[callsign]", UCase(!Callsign))
            subj = Replace(subj, "[type]", UCase(!Type))
            subj = Replace(subj, "[arrdate]", UCase(Format(!arrDate, "d mmm yy")))
            
    '        Dim RS As DAO.Recordset: Set RS = CurrentDb.OpenRecordset(Me.RecordSource)
    '        Dim details As String
    '        With RS: Do While Not .EOF
    '            details = details & vbCrLf & !PPR & " - " & !Callsign & "/" & !acType & " | ARR: " & !arrDate & " / DEP: " & !depDate
    '            .MoveNext
    '        Loop
    '        .Close
    '        End With: Set RS = Nothing
    '        details = details & vbCrLf & vbCrLf
            
            body = DLookup("data", "tblSettings", "key = ""pprEmailBody""")
            body = Replace(body, "[ppr]", !PPR)
            body = Replace(body, "[callsign]", UCase(!Callsign))
            body = Replace(body, "[type]", UCase(!Type))
            body = Replace(body, "[arrdate]", UCase(Format(!arrDate, "d mmm yy")))
            'body = Replace(body, "[details]", UCase(details))
            
            
            'body = body & vbCrLf & vbCrLf & !details
            
            'DoCmd.OpenReport "rPPR", acViewPreview, , "arrDate = #" & arrDate & "#", acHidden, Me.ID
            DoCmd.SendObject acSendReport, "rPPR", acFormatPDF, , , , subj, body, True
            DoCmd.Close acReport, "rPPR"
        End With
        Set RS = Nothing
    End If
    
    
    
    
'
'    Set oEmail = oApp.CreateItem(olMailItem)
'    With oEmail
'        .Subject = Caption
'        .Body = "Attached is the PPR report, current as of " & Format(Now(), "dd-mmm-yy hh:nn L")
'        .Attachments.add file
'    End With

sexit:
    Exit Sub
    Resume Next
errtrap:
    ErrHandler err, Error$, Me.Name
End Sub

Private Sub btnPrint_Click()
On Error Resume Next
    DoCmd.RunCommand acCmdPrint
End Sub

Private Sub depdate_Label_Click()
doOrderDepDate = True
obDepDate = Not obDepDate
refreshOrder 2
End Sub

Public Function doHighlight(ByVal xID As Integer) As Boolean
'This function is called from the 'Conditional Formatting' module(?) of the report
    For Each h In Split(hLight, ",")
        
        doHighlight = (xID = h)
        
        If doHighlight Then Exit Function
    Next
End Function

Private Sub PPR_Label_Click()
doOrderPPR = True
obPPR = Not obPPR
refreshOrder 3
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

'Private Sub Report_Current()
'On Error GoTo errtrap
'    If Not IsNull(Parent.Name) Then
'        'silently fail if parent doesn't exist
'        btnClose.Visible = False
'    End If
'sExit:
'    Exit Sub
'errtrap:
'    Select Case err
'    Case 2452: Resume Next
'    End Select
'    ErrHandler err, Error$, Me.Name & ".Report_Current"
'End Sub

Private Sub Report_Load()
    Caption = "PPR Report " & Format(Now, "dd-mmm-yy hhnn(L)")
    If Nz(OpenArgs) <> "" Then
        Me.RecordSource = OpenArgs
    Else
        Dim newSQL As String
        Dim db As DAO.Database: Set db = CurrentDb
        
        newSQL = Replace(db.QueryDefs("qPPRReport").sql, "[varDate]", "#" & Parent.Parent.reportDate & "#")
        Me.RecordSource = newSQL
        
    End If
End Sub

Private Sub btnOverlay_Click()
Dim idx As Integer
Dim s() As String
Dim found As Boolean

    Dim i: For Each i In Split(Nz(hLight, "0"), ",")
        If ID <> CInt(i) Then
            ReDim Preserve s(0 To idx) As String
            s(idx) = i
            idx = idx + 1
        Else
            found = True
        End If
    Next
    If Not found Then
        ReDim Preserve s(0 To idx) As String
        s(idx) = ID
    End If
    
    hLight = join(s, ",")
    Report.Requery
    Me.btnEmail.SetFocus
End Sub
