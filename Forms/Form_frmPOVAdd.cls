VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmPOVAdd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Dim bClose As Boolean

Private Sub delete_Click()
Dim rs As DAO.Recordset
If MsgBox("This action cannot be undone. Continue?", vbQuestion + vbYesNo, "POV Pass") = vbNo Then Exit Sub
sending.Caption = "Sending to SharePoint..."
sending.Visible = True
DoEvents
Set rs = Me.RecordsetClone
rs.delete
sending.Caption = "Success!"
DoEvents
DoCmd.Close acForm, "frmPOVAdd", acSaveYes

    
End Sub

Private Sub Form_BeforeUpdate(cancel As Integer)
    If Not bClose Then
        'cancel = true
        Undo
        Exit Sub
    End If
End Sub

Private Sub Form_Current()
bClose = False
sending.Caption = "Loading..."
sending.Visible = True
DoEvents
If Nz(Me.Pass_Number) = "" Then Me.Pass_Number = getPassNumber(IIf(passType = 1, "M", "C"))

    If NewRecord Then
         Me.Pass_Number = getPassNumber(IIf(passType = 1, "M", "C"))
        Status = "Requested"
    Else: Select Case Left(Me.Pass_Number, 1)
        Case "M"
            passType = 1
        Case "C"
            passType = 2
        End Select
    
        If Nz(Me.Authorizations_Restrictions) = "" Then Me.Authorizations_Restrictions = Me.Requested_Areas
        
        For Each ctl In Me.Controls
            If TypeOf ctl Is TextBox Then
                ctl.Value = UCase(ctl.Value)
            End If
        Next
        
    End If
    
    rStatus = IIf(Status = "Requested", 1, IIf([Status] = "Approved", 2, 3))

    'passType.Enabled = NewRecord
    delete.Visible = Not NewRecord
    sending.Visible = False
    title.Caption = "POV Pass Request                                              "
    DoEvents
End Sub

Private Sub Form_Load()
    If NewRecord Then
         Me.Pass_Number = getPassNumber(IIf(passType = 1, "M", "C"))
        Status = "Requested"
    End If
End Sub

Private Sub Form_Open(cancel As Integer)
DoEvents
End Sub

Private Sub passType_AfterUpdate()
 Me.Pass_Number = getPassNumber(IIf(passType = 1, "M", "C"))
End Sub

Private Sub rStatus_AfterUpdate()
Dim lName As String
sending.Caption = "Loading..."
sending.Visible = True
DoEvents
sid = DLookup("ID", "ADPMUsers", "[User name] = '" & Environ$("username") & "'")
lName = Nz(DLookup("[tbluserauth].[rank] & ' ' & Left([tbluserauth].[firstname],1) & '. ' & [tbluserauth].[lastname] & '/' & [tbluserauth].[opinitials]", "tblUserAuth", "[spID] = " & _
   sid))
    Select Case rStatus
        Case 1
            Status = "Requested"
        Case 2
            Status = "Approved"
            Me.Date_Issued = Date
            Me.Expiration_Date = DateAdd("yyyy", 1, Me.Date_Issued)
            If Nz(Me.Requested_Areas) = "" Then Me.Authorizations_Restrictions = Me.Requested_Areas
        Case 3
            Status = "Denied"
     End Select
     
     If IsNull(Me.AM_Representitive) Then Me.AM_Representitive = sid
     sending.Visible = False
     DoEvents
    
End Sub

Private Sub save_Click()
On Error GoTo errTrap
bClose = True
sending.Caption = "Sending to SharePoint..."
sending.Visible = True
DoEvents
If Me.Dirty Then Me.Dirty = False

sexit:
    sending.Caption = "Success!"
    DoEvents
    DoCmd.Close acForm, "frmPOVAdd", acSaveYes
    Exit Sub
errTrap:
    MsgBox Error$
End Sub
