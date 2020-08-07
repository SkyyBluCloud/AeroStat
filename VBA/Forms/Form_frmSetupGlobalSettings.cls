VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmSetupGlobalSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub btnBrowse_Click()
Dim fd As Office.FileDialog
Set fd = Access.FileDialog(msoFileDialogFolderPicker)
With fd
    .Filters.clear
    .title = "Select database ROOT location"
    If .show Then
        Dim s: For Each s In .SelectedItems
            dbRoot = s
        Next
    Else
        'Cancelled by user
    End If
End With
End Sub

Private Sub btnChecklists_Click()
DoCmd.OpenForm "frmChecklistsDash"
End Sub

Private Sub btnLogTemplates_Click()
DoCmd.OpenForm "frm3616Template"
End Sub

Private Sub btnRecipAdd_Click()
Dim RS As DAO.Recordset: Set RS = CurrentDb.OpenRecordset("tblPPREmailRecipients")
DoCmd.OpenForm "frmSetupGlobalSettingsPPREmailRecipients", , , , , acDialog
'If Nz(result) = "" Then Exit Sub
Dim result As String: result = Nz(Forms!frmSetupGlobalSettingsPPREmailRecipients!response)
If result = "" Then GoTo sExit
If Right(result, 1) = ";" Then result = Left(result, Len(result) - 1)


If Nz(result) <> "" Then
    Dim s: For Each s In Split(result, ";")
        RS.AddNew
        RS!Recipient = Trim(s)
        RS.Update
    Next
End If
sExit:
    DoCmd.Close acForm, "frmSetupGlobalSettingsPPREmailRecipients"
    lstRecip.Requery
End Sub

Private Sub btnRecipRemove_Click()
On Error Resume Next
CurrentDb.Execute "DELETE FROM tblPPREmailRecipients WHERE ID = " & lstRecip
lstRecip.Requery
End Sub

Private Sub btnShiftDef_Click()
    DoCmd.OpenForm "frmShiftDef"
End Sub

Private Sub Form_AfterUpdate()
On Error Resume Next
    If Parent.Name = "frmBackOffice" Then
        Parent.Form!lblSave.Visible = Parent.Form.Dirty
    End If
End Sub

Private Sub Form_Dirty(cancel As Integer)
On Error Resume Next
    If Parent.Name = "frmBackOffice" Then
        Parent.Form!lblSave.Visible = True
    End If
End Sub

Private Sub Form_Load()
'lstRecip.RowSource = Nz(pprEmailRecipients)
lstRecip = ""
End Sub
