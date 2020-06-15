VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmPOVPass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub bApprove_Click()
On Error GoTo errTrap
If IsNull(requestBoard.Column(0)) Then Exit Sub
If MsgBox("This will mark all selected requests as approved. Continue?", vbQuestion + vbYesNo, "POV Pass") = vbNo Then Exit Sub
Dim pass As String
Dim s() As String
Dim c As String
Dim rs As DAO.Recordset
Dim db As DAO.Database
Dim rc As Integer
Set db = CurrentDb
sending.Caption = "Sending to SharePoint..."
sending.Visible = True
DoEvents
sid = DLookup("ID", "ADPMUsers", "[User name] = '" & Environ$("username") & "'")
    
    Select Case MsgBox("Are these MILITARY Requests?", vbQuestion + vbYesNo, "POV Pass")
    Case vbYes
        pass = "M"
    Case vbNo
        pass = "C"
    End Select

Set rs = db.OpenRecordset(getSelection)
With rs: Do While Not .EOF
    .edit
    rc = rc + 1
    sending.Caption = "Processing " & rc & " of " & .RecordCount & " requests..."
    DoEvents
    !Status = "Approved"
    If Nz(.Fields("Pass Number")) = "" Then .Fields("Pass Number") = getPassNumber(pass)
    .Fields("Date Issued") = Date
    If Nz(.Fields("Authorizations/Restrictions")) = "" Then .Fields("Authorizations/Restrictions") = UCase(.Fields("Requested Areas"))
    .Fields("Expiration Date") = Nz(.Fields("Expiration Date"), DateSerial(Year(.Fields("Date Issued")) + 1, Month(.Fields("Date Issued")), Day(.Fields("Date Issued"))))
    .Fields("AM Representitive") = sid
    If Nz(.Fields("Requested Areas")) = "" Then .Fields("Authorizations/Restrictions") = UCase(.Fields("Requested Areas"))
    'On Error Resume Next
    i = 0
    For Each fld In .Fields
        i = i + 1
        If i = 22 Then Exit For
        fld = UCase(fld)
    Next
    'On Error GoTo errtrap
    .update
    .MoveNext
Loop: End With

sending.Caption = "Success!"
DoEvents
requestBoard.Requery
sending.Visible = False

sexit:
    Exit Sub
errTrap:
    errHandler err, Error$, Name
    sending.Caption = "Unable to process: " & Error$ & " (" & err & ")"
    DoEvents
    Exit Sub
End Sub

Private Function getSelection() As String
Dim s() As String
Dim c As String
Dim idx As Integer

    For Each i In requestBoard.ItemsSelected
        If requestBoard.Column(2, i) = "Requested" Then
            ReDim Preserve s(0 To idx) As String
            s(idx) = requestBoard.Column(0, i)
            idx = idx + 1
        End If
    Next
    
    On Error Resume Next
    If LBound(s) = 0 Then
        getSelection = "SELECT * FROM qPOV WHERE qPOV.ID = " & join(s, " OR qPOV.ID = ")
    End If
    
End Function


Private Sub bNew_Click()
DoCmd.OpenForm "frmPOVAdd", , , , , acDialog
requestBoard.Requery
End Sub

Private Sub bPrint_Click()
On Error GoTo errTrap
If IsNull(requestBoard.Column(0)) Then Exit Sub
Dim s() As String
Dim c As String
Dim idx As Integer

    For Each i In requestBoard.ItemsSelected
        If requestBoard.Column(2, i) = "Approved" Then
            ReDim Preserve s(0 To idx) As String
            s(idx) = requestBoard.Column(0, i)
            idx = idx + 1
        End If
    Next
    
    On Error GoTo printfail
    If UBound(s) <> 0 Or idx = 1 Then
        c = "qPOV.ID = " & join(s, " OR qPOV.ID = ")
        DoCmd.OpenReport "rPOVPass", acViewReport, , c
    End If
sexit:
    Exit Sub
errTrap:
    errHandler err, Error$, Name
    Exit Sub
printfail:
    MsgBox "There were no approved passes selected. Print failed.", vbInformation, "POV Pass"
End Sub

Private Sub requestBoard_DblClick(cancel As Integer)
If requestBoard.ItemsSelected.Count <> 1 Then Exit Sub
sending.Caption = "Opening POV Pass..."
sending.Visible = True
DoEvents
DoCmd.OpenForm "frmPOVAdd", , , "qPOV.ID = " & requestBoard.Column(0), acFormEdit, acDialog
sending.Visible = False
requestBoard.Requery
End Sub
